import { spawn } from 'node:child_process'
import { appendFileSync, mkdirSync, statSync, unlinkSync } from 'node:fs'
import { dirname, join } from 'node:path'

const DEFAULTS = {
  /** 通知发送者身份；DeepSeek Harness 安装包注册的值，换安装渠道时才需要改。 */
  appId: 'com.deepseek.dsh',
  /** Toast 第一行：程序名。 */
  title: 'DeepSeek Harness',
  /** 'auto' | 'zh' | 'en'。 */
  language: 'auto',
  /** 是否播放通知提示音。 */
  sound: true,
  /** 决策类通知是否用 Windows「提醒」场景常驻，直到用户处理。 */
  persistentDecisions: true,
  notifyApproval: true,
  notifyQuestion: true,
  notifyTurnEnd: true,
  notifyError: true,
  notifyBalance: true,
  notifyAccount: true,
  /** 子代理（subagent）的完成/出错是否也提醒。 */
  notifySubagents: false,
  /** 正文截断长度（字符）。 */
  bodyLimit: 220,
  /** 同一条消息在这个窗口内只弹一次，避免重试风暴刷屏。 */
  dedupeWindowMs: 4000,
  /** 自定义 powershell.exe 路径；空字符串表示自动探测。 */
  powershell: '',
  /** 是否写诊断日志（Toast 是否真的弹出、失败原因都记在这里）。 */
  log: true,
  /** 诊断日志路径；空字符串表示 %LOCALAPPDATA%\\DeepSeekHarnessToast\\notifications.log。 */
  logFile: '',
}

const TEXT = {
  zh: {
    approval: detail => `需要你的确认：${detail}`,
    approvalFallback: tool => `即将执行工具 ${tool}`,
    question: text => `需要你的选择：${text}`,
    questionMore: count => `（还有 ${count} 个问题）`,
    turnDone: '任务已完成，等你查看结果。',
    turnGoal: '本轮已完成，目标仍在自动继续。',
    error: message => `任务出错：${message}`,
    balance: '余额不足：DeepSeek 账户余额不足，请充值后再继续。',
    signIn: '需要登录 DeepSeek 账号后才能继续。',
    expired: '登录已过期，请重新登录 DeepSeek 账号。',
  },
  en: {
    approval: detail => `Needs your approval: ${detail}`,
    approvalFallback: tool => `about to run ${tool}`,
    question: text => `Needs your answer: ${text}`,
    questionMore: count => ` (and ${count} more)`,
    turnDone: 'Task finished. Check the result.',
    turnGoal: 'This round finished; the goal continues automatically.',
    error: message => `Task failed: ${message}`,
    balance: 'Insufficient balance: top up your DeepSeek account to continue.',
    signIn: 'Sign in to your DeepSeek account to continue.',
    expired: 'Your DeepSeek sign-in expired. Please sign in again.',
  },
}

const BALANCE_PATTERN = /insufficient[_\s-]?balance|insufficient[_\s-]?quota|余额不足|payment[_\s-]?required|credit[_\s-]?balance/i
const ABORT_PATTERN = /abort|cancel|interrupt|terminated by user|用户(取消|中断)/i

const POWERSHELL_SCRIPT = `
$ErrorActionPreference = 'Stop'
$payload = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PAYLOAD__')) | ConvertFrom-Json
$logPath = [string]$payload.log
function Write-ToastLog([string]$text) {
  if (-not $logPath) { return }
  try {
    $dir = [System.IO.Path]::GetDirectoryName($logPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    if ((Test-Path -LiteralPath $logPath) -and ((Get-Item -LiteralPath $logPath).Length -gt 262144)) { Remove-Item -LiteralPath $logPath -Force }
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [System.IO.File]::AppendAllText($logPath, ($stamp + ' ' + $text + [Environment]::NewLine), [Text.Encoding]::UTF8)
  } catch { }
}
try {
  [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
  [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
  $document = New-Object Windows.Data.Xml.Dom.XmlDocument
  $document.LoadXml([string]$payload.xml)
  $toast = New-Object Windows.UI.Notifications.ToastNotification $document
  [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier([string]$payload.appId).Show($toast)
  Write-ToastLog ('ok app=' + $payload.appId + ' body=' + $payload.body)
  exit 0
} catch {
  Write-ToastLog ('fail ' + $_.Exception.GetType().Name + ': ' + $_.Exception.Message)
  exit 1
}
`

/** 读取配置，未知键忽略，类型不符时回落到默认值。 */
function resolveConfig(raw) {
  const config = { ...DEFAULTS }
  const source = raw && typeof raw === 'object' ? raw : {}
  for (const key of Object.keys(DEFAULTS)) {
    const value = source[key]
    if (value === undefined || value === null) continue
    const expected = typeof DEFAULTS[key]
    if (typeof value === expected) config[key] = value
  }
  if (typeof source.language === 'string' && ['auto', 'zh', 'en'].includes(source.language)) {
    config.language = source.language
  }
  if (!Number.isFinite(config.dedupeWindowMs) || config.dedupeWindowMs < 0) config.dedupeWindowMs = DEFAULTS.dedupeWindowMs
  if (!Number.isFinite(config.bodyLimit) || config.bodyLimit < 20) config.bodyLimit = DEFAULTS.bodyLimit
  return config
}

function pickLanguage(configured) {
  if (configured === 'zh' || configured === 'en') return configured
  let locale = ''
  try {
    locale = Intl.DateTimeFormat().resolvedOptions().locale || ''
  } catch {
    locale = ''
  }
  return locale.toLowerCase().startsWith('zh') ? 'zh' : 'en'
}

function defaultLogFile() {
  const base = process.env.LOCALAPPDATA || process.env.XDG_STATE_HOME || ''
  if (!base) return ''
  return join(base, 'DeepSeekHarnessToast', 'notifications.log')
}

function defaultPowerShell() {
  const root = process.env.SystemRoot || 'C:\\Windows'
  return join(root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
}

function escapeXml(text) {
  return String(text)
    // XML 1.0 不允许的控制字符
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function truncate(text, limit) {
  const value = String(text ?? '').replace(/\s+/g, ' ').trim()
  const characters = Array.from(value)
  if (characters.length <= limit) return value
  return `${characters.slice(0, Math.max(1, limit - 1)).join('')}…`
}

function buildToastXml({ title, lines, sound, persistent }) {
  // ToastGeneric 最多显示 3 行：标题 + 两行正文。
  const texts = [title, ...lines.filter(Boolean).slice(0, 2)].map(line => `<text>${escapeXml(line)}</text>`).join('')
  const scenario = persistent ? ' scenario="reminder"' : ''
  const audio = sound ? (persistent ? '<audio/>' : '') : '<audio silent="true"/>'
  return `<toast${scenario}><visual><binding template="ToastGeneric">${texts}</binding></visual>${audio}</toast>`
}

function errorMessageOf(error) {
  if (error === undefined || error === null) return ''
  if (typeof error === 'string') return error
  if (typeof error.message === 'string') return error.message
  try {
    return String(error)
  } catch {
    return ''
  }
}

/** 诊断日志：Toast 是否真的发出、PowerShell 为何失败，都写在这里。 */
function writeDiagnostic(logFile, line) {
  if (!logFile) return
  try {
    mkdirSync(dirname(logFile), { recursive: true })
    try {
      if (statSync(logFile).size > 262144) unlinkSync(logFile)
    } catch {
      // 文件不存在或不可读时忽略
    }
    appendFileSync(logFile, `${new Date().toISOString()} [host] ${line}\n`, 'utf8')
  } catch {
    // 日志不可写不影响通知
  }
}

export function apply(ctx, rawConfig) {
  if (process.platform !== 'win32') return

  const config = resolveConfig(rawConfig)
  const language = pickLanguage(config.language)
  const text = TEXT[language]
  const logFile = config.log ? (config.logFile || defaultLogFile()) : ''
  const powershell = config.powershell || defaultPowerShell()
  const seen = new Map()

  writeDiagnostic(logFile, `started app=${config.appId} language=${language} powershell=${powershell}`)

  const warn = message => {
    try {
      console.warn(`[dsh-toast] ${message}`)
    } catch {
      // 忽略日志失败
    }
  }

  /** 同一 key 在窗口期内只放行一次；命中时也刷新时间戳，压住重试风暴。 */
  const throttled = (key, windowMs = config.dedupeWindowMs) => {
    const now = Date.now()
    const last = seen.get(key)
    if (last !== undefined && now - last < windowMs) {
      seen.set(key, now)
      return true
    }
    seen.set(key, now)
    if (seen.size > 96) {
      for (const [entry, at] of seen) {
        if (now - at > 60000) seen.delete(entry)
      }
    }
    return false
  }

  const showToast = ({ lines, persistent = false, sound = config.sound, category }) => {
    const body = lines.find(line => !!line) || ''
    const xml = buildToastXml({ title: config.title, lines, sound, persistent })
    const payload = { appId: config.appId, xml, body: truncate(body, 120), log: logFile }
    const encodedPayload = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64')
    const script = POWERSHELL_SCRIPT.replace('__PAYLOAD__', encodedPayload)
    const encodedScript = Buffer.from(script, 'utf16le').toString('base64')
    try {
      const child = spawn(
        powershell,
        ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-WindowStyle', 'Hidden', '-EncodedCommand', encodedScript],
        { windowsHide: true, stdio: 'ignore' },
      )
      child.on('error', error => {
        const message = `无法启动 PowerShell（${category}）：${errorMessageOf(error)}`
        warn(message)
        writeDiagnostic(logFile, message)
      })
      child.unref()
    } catch (error) {
      const message = `Toast 发送失败（${category}）：${errorMessageOf(error)}`
      warn(message)
      writeDiagnostic(logFile, message)
    }
  }

  const sessionOf = agent => {
    try {
      return ctx.get('sessions')?.get(agent?.id)
    } catch {
      return undefined
    }
  }

  const sessionTitleOf = session => {
    if (!session) return ''
    try {
      return truncate(ctx.get('sessionTitle')?.get(session)?.title ?? '', 48)
    } catch {
      return ''
    }
  }

  const isRootAgent = agent => {
    const session = sessionOf(agent)
    if (!session) return true
    const header = session.header ?? {}
    return !header.parentSession && header.origin !== 'subagent'
  }

  const activeGoal = agent => {
    try {
      const goal = ctx.get('goals')?.get(agent)
      return goal && goal.phase === 'active' ? goal : undefined
    } catch {
      return undefined
    }
  }

  const notify = ({ category, lines, agent, persistent = false, sound = config.sound, key, windowMs }) => {
    const session = sessionOf(agent)
    const sessionTitle = sessionTitleOf(session)
    const content = sessionTitle ? [...lines, sessionTitle] : lines
    if (throttled(key, windowMs)) return
    showToast({ category, lines: content, persistent, sound })
  }

  // ── 需要用户决策：工具执行授权 ────────────────────────────────────────────
  ctx.on('approval/request', (request, next) => {
    try {
      if (config.notifyApproval) {
        const reason = request?.displayReason?.[language] ?? request?.displayReason?.en ?? request?.reason
        const tool = request?.toolName ? String(request.toolName) : ''
        const detail = truncate(reason || text.approvalFallback(tool || 'tool'), 160)
        notify({
          category: 'approval',
          agent: request?.agent,
          persistent: config.persistentDecisions,
          lines: [text.approval(detail), tool ? `工具：${tool}` : ''].filter(Boolean),
          key: `approval|${request?.agent?.id ?? ''}|${detail}`,
          windowMs: 2000,
        })
      }
    } catch (error) {
      warn(`授权提醒失败：${errorMessageOf(error)}`)
    }
    return next()
  })

  // ── 需要用户决策：结构化提问 ─────────────────────────────────────────────
  ctx.on('user-questions/request', (request, next) => {
    try {
      if (config.notifyQuestion) {
        const questions = Array.isArray(request?.questions) ? request.questions : []
        const first = questions[0]
        const head = first?.header || first?.question || ''
        const more = questions.length > 1 ? text.questionMore(questions.length - 1) : ''
        notify({
          category: 'question',
          agent: request?.agent,
          persistent: config.persistentDecisions,
          lines: [text.question(truncate(head, 160) + more)],
          key: `question|${request?.agent?.id ?? ''}|${head}`,
          windowMs: 2000,
        })
      }
    } catch (error) {
      warn(`提问提醒失败：${errorMessageOf(error)}`)
    }
    return next()
  })

  // ── 任务结束 ─────────────────────────────────────────────────────────────
  ctx.on('agent/turn-stopping', payload => {
    try {
      const agent = payload?.agent
      if (!config.notifyTurnEnd) return
      if (!agent) return
      if (!config.notifySubagents && !isRootAgent(agent)) return
      const goal = activeGoal(agent)
      notify({
        category: 'turn',
        agent,
        lines: [goal ? text.turnGoal : text.turnDone],
        key: `turn|${agent.id}`,
        windowMs: 1500,
      })
    } catch (error) {
      warn(`任务结束提醒失败：${errorMessageOf(error)}`)
    }
  })

  // ── 余额不足 / 其它模型请求失败 ──────────────────────────────────────────
  ctx.on('agent/request-error', async (payload, next) => {
    try {
      const failure = payload?.failure
      const agent = payload?.agent
      if (failure && config.notifyBalance && (failure.status === 402 || BALANCE_PATTERN.test(`${failure.code ?? ''} ${failure.message ?? ''}`))) {
        notify({
          category: 'balance',
          agent,
          lines: [text.balance, truncate(failure.message ?? '', 120)],
          key: `balance|${agent?.id ?? ''}`,
          windowMs: 60000,
        })
      }
    } catch (error) {
      warn(`余额提醒失败：${errorMessageOf(error)}`)
    }
    return next()
  })

  // ── 任务出错 ─────────────────────────────────────────────────────────────
  ctx.on('agent/error', payload => {
    try {
      if (!config.notifyError) return
      const agent = payload?.agent
      if (!config.notifySubagents && agent && !isRootAgent(agent)) return
      const error = payload?.error
      const message = truncate(errorMessageOf(error), 140)
      if (!message || ABORT_PATTERN.test(message)) return
      // 余额不足已由 agent/request-error 单独提醒，避免同一件事弹两条。
      if (config.notifyBalance && error && typeof error === 'object') {
        if (error.status === 402 || BALANCE_PATTERN.test(`${error.code ?? ''} ${error.message ?? ''}`)) return
      }
      notify({
        category: 'error',
        agent,
        lines: [text.error(message)],
        key: `error|${agent?.id ?? ''}|${message}`,
        windowMs: 15000,
      })
    } catch (error) {
      warn(`出错提醒失败：${errorMessageOf(error)}`)
    }
  })

  // ── 账号需要登录 / 登录失效 ──────────────────────────────────────────────
  ctx.on('deepseek-account/model-sign-in-required', () => {
    try {
      if (!config.notifyAccount) return
      if (throttled('account|signin', 60000)) return
      showToast({ category: 'account', lines: [text.signIn] })
    } catch (error) {
      warn(`登录提醒失败：${errorMessageOf(error)}`)
    }
  })

  ctx.on('deepseek-account/session-expired', () => {
    try {
      if (!config.notifyAccount) return
      if (throttled('account|expired', 60000)) return
      showToast({ category: 'account', lines: [text.expired] })
    } catch (error) {
      warn(`登录失效提醒失败：${errorMessageOf(error)}`)
    }
  })
}
