/**
 * dsh-doc-review client half (v0.1)
 *
 * Sidebar annotations over the built-in document preview:
 * - Word/PDF: select text in the preview, add a note, collect several, submit once.
 * - Excel: add a note anchored to a worksheet cell address (manual in v0.1).
 *
 * The client half uses only the module-loader contract: it registers into
 * `sidebar.right.tab.document.actions` and sends one prompt through the
 * session binding. No Host service or build step is required.
 */
window.__ModuleLoader__.load({
  id: 'dsh-doc-review',
  factory: (require) => {
    const module = { exports: {} }
    const exports = module.exports
    const React = require('react')
    const h = React.createElement
    const { useEffect, useRef, useState } = React

    const STYLE_ID = 'dsh-doc-review-style'
    const HIGHLIGHT = 'dsh-doc-review'
    const SHEET_RE = /\.(xlsx|xlsm|xls|csv|tsv)$/i
    const TEXT_RE = /\.(docx?|pdf)$/i

    const CSS_TEXT = `
.dr-btn{display:inline-flex;align-items:center;gap:5px;height:24px;padding:0 8px;border:1px solid var(--dsw-alias-border-l1);border-radius:6px;background:transparent;color:var(--dsw-alias-label-primary);font-size:12px;line-height:1;cursor:pointer}
.dr-btn:hover{background:var(--dsw-alias-bg-layer-2)}
.dr-btn-on{border-color:var(--dsw-alias-brand-primary);color:var(--dsw-alias-brand-primary)}
.dr-count{min-width:16px;height:16px;padding:0 4px;border-radius:8px;background:var(--dsw-alias-brand-primary);color:#fff;font-size:10px;line-height:16px;text-align:center}
.dr-panel{position:fixed;right:16px;bottom:88px;z-index:2147483000;display:flex;flex-direction:column;width:340px;max-width:calc(100vw - 32px);max-height:60vh;border:1px solid var(--dsw-alias-border-l2);border-radius:10px;background:var(--dsw-alias-bg-overlay);color:var(--dsw-alias-label-primary);box-shadow:0 8px 32px rgba(0,0,0,.18);font-size:12px}
.dr-head{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid var(--dsw-alias-border-l1)}
.dr-head strong{font-size:13px}
.dr-sub{flex:1;color:var(--dsw-alias-label-secondary)}
.dr-x{border:0;background:transparent;color:var(--dsw-alias-label-secondary);font-size:16px;line-height:1;cursor:pointer}
.dr-hint{margin:10px 12px;color:var(--dsw-alias-label-secondary);line-height:1.5}
.dr-tip{margin:0 12px 8px;color:var(--dsw-alias-label-secondary);font-size:11px;line-height:1.45}
.dr-list{margin:0;padding:0 12px;list-style:none;overflow:auto}
.dr-item{padding:8px 0;border-top:1px solid var(--dsw-alias-border-l1)}
.dr-item:first-child{border-top:0}
.dr-meta{color:var(--dsw-alias-label-secondary);font-size:11px;margin-bottom:4px}
.dr-quote{margin:0 0 5px;padding-left:8px;border-left:2px solid var(--dsw-alias-border-l2);color:var(--dsw-alias-label-secondary);line-height:1.45;max-height:72px;overflow:auto}
.dr-note{margin:0 0 4px;line-height:1.5;white-space:pre-wrap}
.dr-del{border:0;background:transparent;color:var(--dsw-alias-label-secondary);font-size:11px;cursor:pointer;padding:0}
.dr-del:hover{color:var(--dsw-alias-state-error-primary)}
.dr-foot{display:flex;align-items:center;justify-content:flex-end;gap:8px;padding:10px 12px;border-top:1px solid var(--dsw-alias-border-l1)}
.dr-primary{height:28px;padding:0 12px;border:0;border-radius:6px;background:var(--dsw-alias-brand-primary);color:#fff;font-size:12px;cursor:pointer}
.dr-primary:disabled{opacity:.45;cursor:default}
.dr-error{margin:8px 12px 0;color:var(--dsw-alias-state-error-primary);line-height:1.45}
.dr-status{margin:8px 12px 0;color:var(--dsw-alias-state-success-primary);line-height:1.45}
.dr-bubble{position:fixed;z-index:2147483001;height:26px;padding:0 10px;border:0;border-radius:6px;background:var(--dsw-alias-brand-primary);color:#fff;font-size:12px;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.dr-pop{position:fixed;z-index:2147483001;width:300px;max-width:calc(100vw - 16px);padding:10px;border:1px solid var(--dsw-alias-border-l2);border-radius:10px;background:var(--dsw-alias-bg-overlay);color:var(--dsw-alias-label-primary);box-shadow:0 8px 28px rgba(0,0,0,.2);font-size:12px}
.dr-pop textarea,.dr-form textarea,.dr-form input{box-sizing:border-box;width:100%;margin:0 0 6px;padding:6px 8px;border:1px solid var(--dsw-alias-border-l1);border-radius:6px;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font:inherit;font-size:12px}
.dr-pop textarea,.dr-form textarea{min-height:56px;resize:vertical}
.dr-actions{display:flex;justify-content:flex-end;gap:8px}
.dr-actions button{height:26px;padding:0 10px;border:1px solid var(--dsw-alias-border-l1);border-radius:6px;background:transparent;color:var(--dsw-alias-label-primary);font-size:12px;cursor:pointer}
.dr-actions .dr-primary{border:0;background:var(--dsw-alias-brand-primary);color:#fff}
.dr-form{display:flex;flex-direction:column;padding:10px 12px 0}
.dr-form .dr-row{display:flex;gap:6px}
.dr-form .dr-row input{flex:1}
.dr-form button{height:28px;margin-bottom:8px;border:1px solid var(--dsw-alias-border-l1);border-radius:6px;background:transparent;color:var(--dsw-alias-label-primary);font-size:12px;cursor:pointer}
::highlight(dsh-doc-review){background-color:rgba(255,196,0,.35)}
`

    function ensureStyle() {
      if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return
      const style = document.createElement('style')
      style.id = STYLE_ID
      style.textContent = CSS_TEXT
      document.head.appendChild(style)
    }

    function uid() {
      try {
        return crypto.randomUUID()
      } catch {
        return `${Date.now()}-${Math.random().toString(16).slice(2)}`
      }
    }


    const STORAGE_PREFIX = 'dsh-doc-review:v1:'

    function loadStored(path) {
      try {
        const raw = localStorage.getItem(STORAGE_PREFIX + path)
        const list = raw ? JSON.parse(raw) : []
        return Array.isArray(list) ? list.filter(item => item && typeof item === 'object') : []
      } catch {
        return []
      }
    }

    function saveStored(path, comments) {
      try {
        const serial = comments.map(({ range, ...rest }) => rest)
        localStorage.setItem(STORAGE_PREFIX + path, JSON.stringify(serial))
      } catch {
        /* quota or private mode */
      }
    }

    function clearStored(path) {
      try {
        localStorage.removeItem(STORAGE_PREFIX + path)
      } catch {
        /* ignore */
      }
    }

    function collectText(root) {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      const nodes = []
      let text = ''
      let node
      while ((node = walker.nextNode())) {
        const value = node.nodeValue || ''
        if (!value) continue
        nodes.push({ node, start: text.length, end: text.length + value.length })
        text += value
      }
      return { nodes, text }
    }

    function normalizeMap(text) {
      let norm = ''
      const map = []
      for (let index = 0; index < text.length; index += 1) {
        const char = text[index]
        if (!/\s/.test(char)) {
          norm += char
          map.push(index)
        }
      }
      return { norm, map }
    }

    function findRangeIn(root, quote) {
      if (!root || !quote) return null
      const target = normalizeMap(quote).norm
      if (!target) return null
      const { nodes, text } = collectText(root)
      const { norm, map } = normalizeMap(text)
      const startIndex = norm.indexOf(target)
      if (startIndex === -1) return null
      const start = map[startIndex]
      const end = map[startIndex + target.length - 1] + 1
      let startNode = null
      let startOffset = 0
      let endNode = null
      let endOffset = 0
      for (const item of nodes) {
        if (!startNode && start >= item.start && start < item.end) {
          startNode = item.node
          startOffset = start - item.start
        }
        if (end > item.start && end <= item.end) {
          endNode = item.node
          endOffset = end - item.start
          break
        }
      }
      if (!startNode || !endNode) return null
      try {
        const range = document.createRange()
        range.setStart(startNode, startOffset)
        range.setEnd(endNode, endOffset)
        return range
      } catch {
        return null
      }
    }

    function findRangeForComment(comment) {
      const root = comment.page
        ? document.querySelector(`[data-pdf-page="${comment.page}"]`)
        : document.querySelector('[data-pdf-preview]') || document.querySelector('[data-textpreview-body]')
      return findRangeIn(root, comment.quote)
    }

    function closestElement(node, selector) {
      const el = node && node.nodeType === 3 ? node.parentElement : node
      return el && el.closest ? el.closest(selector) : null
    }

    function readSelection() {
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null
      const range = sel.getRangeAt(0)
      if (closestElement(range.commonAncestorContainer, '[data-doc-review]')) return null
      const body = closestElement(range.commonAncestorContainer, '[data-pdf-preview], [data-textpreview-body]')
      if (!body) return null
      const quote = sel.toString().replace(/\s+/g, ' ').trim().slice(0, 500)
      if (!quote) return null
      const pageEl = closestElement(range.startContainer, '[data-pdf-page], [data-page-number]')
      const page = pageEl
        ? pageEl.getAttribute('data-pdf-page') || pageEl.getAttribute('data-page-number') || ''
        : ''
      let rect
      try {
        rect = range.getBoundingClientRect()
      } catch {
        return null
      }
      if (!rect || (!rect.width && !rect.height)) return null
      return { range: range.cloneRange(), quote, page, x: rect.left, y: rect.bottom }
    }

    function buildMessage(path, comments) {
      const lines = [
        '【批注修改任务 · dsh-doc-review v1.5】',
        `文件：${path}`,
        `批注：${comments.length} 条`,
        '预览：提交时已在右侧栏打开，修改后自动刷新；完成后不要调用 present，报告里给出文件路径即可。',
        '',
      ]
      comments.forEach((item, index) => {
        const no = index + 1
        if (item.cell) {
          lines.push(`${no}. [${item.sheet || 'Sheet1'}!${item.cell}]`)
          if (item.value) lines.push(`   现值：“${item.value}”`)
          lines.push(`   意见：${item.note}`)
        } else {
          lines.push(`${no}. [第 ${item.page || '?'} 页]`)
          lines.push(`   引文：“${item.quote}”`)
          lines.push(`   意见：${item.note}`)
        }
        lines.push('')
      })
      lines.push(
        '执行约定（仅约束本次文件修改的操作方式；内容与文风仍遵循当前模式，安全、权限与审批遵循宿主）：',
        '0. 需求已由批注确认，不重复提纲、方案确认或先例检索，直接修改。',
        '1. 只改批注涉及的内容，不顺手优化其他段落；只读目标文件，不浏览工作区其他文件，不检索记忆或知识库。',
        '2. 第一步就是运行 `~/.dsh/doc-review/bin/docx_edit.py`：--file 目标文件，--ops 传内联 JSON 或 ops 文件（op 支持 replace / delete / set_font / set_color / set_paragraph_style，可用 all 和 occurrence）。它内部完成 hash、备份、扫描、修改、回读；在此之前不运行 ls、cat 或单独的检查命令。不要再写临时脚本；只有 helper 返回 unsupported 时才写一次性脚本。',
        '3. helper 只输出摘要：hash、备份名、命中段落号、改动摘要、校验结果；不要把全部段落或 XML 打印出来。',
        '4. 先写临时文件，全部断言通过后再原子替换；中途失败原文件不变。',
        '5. 引文先去空白再匹配；多命中按页码与上下文择一，仍不唯一则报“未找到/有歧义”，不猜。',
        '6. 保留原样式、字体、编号、表格、页眉页脚；新增文字与上下文风格一致，不主动改文风。',
        '7. Excel 按 sheet!cell + 现值 + 相邻单元格定位；保留公式与样式，不重算、不导出。',
        '8. 不做渲染验证（不启动 WPS/qlmanage/textutil）；用回读断言校验。',
        '9. 措辞歧义按最小改动处理并说明；事实、口径、范围、权限不清才提问。',
        '10. 目标已是所需状态时报告“已是目标状态”，不重复修改。',
        '11. 沉淀动作只在当前模式明确要求时，于修改与校验完成后一次完成；不因此增加修改轮次。',
        '12. 除 helper、失败日志、最终 present 与报告外，不调用其他工具。',
        '13. helper 报 unsupported，或改为一次性脚本、一次性脚本失败、需要人工介入时，运行 `~/.dsh/doc-review/bin/doc_review_log.py log --event fallback_script|manual_intervention|verify_fail --reason "…" --file "…"` 记录；不写文档全文，引文截断 120 字。helper 自身失败由 helper 自动记录，不重复。',
        '14. 消息声明“预览已打开”时，完成修改后不要调用 present；只有预览未打开或用户明确要交付卡片时才 present。',
        '',
        '报告格式（紧凑，不超过 10 行）：文件 / 备份 / 逐条状态（已改·未找到·需确认 + 一句说明）/ 校验结果。',
      )
      return lines.join('\n')
    }

    async function sendPrompt(ctx, sessionId, text) {
      if (!sessionId) throw new Error('当前没有可用的会话')
      const reference = ctx.sessions.retain(sessionId, { source: 'docReview' })
      try {
        const binding = await reference.ready
        const result = await binding.session.prompt([{ type: 'text', text }], 'queue')
        if (!result || result.ok !== true) {
          const error = result && result.error
          throw new Error(error ? `${error.code}: ${error.message}` : '发送失败')
        }
      } finally {
        reference.release()
      }
    }

    function ReviewAction(props) {
      const path = props.absolutePath
      const isSheet = SHEET_RE.test(path)
      const isText = TEXT_RE.test(path)
      const [open, setOpen] = useState(false)
      const [enabled, setEnabled] = useState(false)
      const [comments, setComments] = useState([])
      const [pending, setPending] = useState(null)
      const [draft, setDraft] = useState('')
      const [busy, setBusy] = useState(false)
      const [error, setError] = useState('')
      const [status, setStatus] = useState('')
      const [cell, setCell] = useState({ sheet: 'Sheet1', addr: '', value: '', note: '' })
      const sendRef = useRef(props.send)
      sendRef.current = props.send

      useEffect(() => {
        const stored = loadStored(path)
        setComments(stored.map(item => ({ ...item, id: item.id || uid() })))
        setPending(null)
        setDraft('')
        setEnabled(false)
        setOpen(false)
        setError('')
        setStatus('')
      }, [path])

      useEffect(() => {
        if (!status) return undefined
        const timer = setTimeout(() => setStatus(''), 4000)
        return () => clearTimeout(timer)
      }, [status])

      useEffect(() => {
        if (!enabled || !isText) return undefined
        let timer = null
        const onChange = () => {
          if (timer) clearTimeout(timer)
          timer = setTimeout(() => {
            const info = readSelection()
            if (info) setPending({ ...info, editing: false })
          }, 120)
        }
        document.addEventListener('selectionchange', onChange)
        return () => {
          if (timer) clearTimeout(timer)
          document.removeEventListener('selectionchange', onChange)
        }
      }, [enabled, isText])

      useEffect(() => {
        if (typeof Highlight !== 'function' || typeof CSS === 'undefined' || !CSS.highlights) return undefined
        const ranges = comments.map(item => item.range).filter(Boolean)
        try {
          if (ranges.length) CSS.highlights.set(HIGHLIGHT, new Highlight(...ranges))
          else CSS.highlights.delete(HIGHLIGHT)
        } catch {
          /* detached ranges after a preview reload */
        }
        return () => {
          try {
            CSS.highlights.delete(HIGHLIGHT)
          } catch {
            /* ignore */
          }
        }
      }, [comments])

      useEffect(() => {
        if (!comments.some(item => item.quote && !item.range)) return undefined
        let tries = 0
        const timer = setInterval(() => {
          setComments(list => {
            let changed = false
            const next = list.map(item => {
              if (item.quote && !item.range) {
                const range = findRangeForComment(item)
                if (range) {
                  changed = true
                  return { ...item, range }
                }
              }
              return item
            })
            return changed ? next : list
          })
          tries += 1
          if (tries >= 12) clearInterval(timer)
        }, 400)
        return () => clearInterval(timer)
      }, [comments])

      const count = comments.length

      if (!isText && !isSheet) return null

      function toggle() {
        const next = !open
        setOpen(next)
        setEnabled(next && isText)
        if (!next) setPending(null)
        else if (isText) setComments(list => list.map(({ range, ...rest }) => rest))
      }

      function commit(next) {
        setComments(next)
        saveStored(path, next)
      }

      function saveText() {
        const note = draft.trim()
        if (!pending || !note) return
        commit([...comments, { id: uid(), quote: pending.quote, page: pending.page, note, range: pending.range }])
        setPending(null)
        setDraft('')
      }

      function addCell() {
        const addr = cell.addr.trim().toUpperCase()
        if (!addr || !cell.note.trim()) return
        commit([
          ...comments,
          {
            id: uid(),
            sheet: cell.sheet.trim() || 'Sheet1',
            cell: addr,
            value: cell.value.trim(),
            note: cell.note.trim(),
          },
        ])
        setCell({ sheet: cell.sheet, addr: '', value: '', note: '' })
      }

      async function submit() {
        if (!count || busy) return
        const send = sendRef.current
        if (!send) {
          setError('发送通道不可用')
          return
        }
        setBusy(true)
        setError('')
        try {
          await send(buildMessage(path, comments))
          clearStored(path)
          setComments([])
          setEnabled(false)
          setOpen(false)
          setStatus(`已提交 ${count} 条批注`)
        } catch (cause) {
          setError(cause && cause.message ? cause.message : String(cause))
        } finally {
          setBusy(false)
        }
      }

      const button = h(
        'button',
        {
          className: 'dr-btn' + (open ? ' dr-btn-on' : ''),
          title: isSheet ? '按单元格添加批注' : '在预览中选中文字添加批注',
          onClick: toggle,
        },
        '批注',
        count ? h('span', { className: 'dr-count' }, String(count)) : null,
      )

      const panel = open
        ? h(
            'section',
            { className: 'dr-panel', 'data-doc-review': 'panel' },
            h(
              'header',
              { className: 'dr-head' },
              h('strong', null, '评审批注'),
              h('span', { className: 'dr-sub' }, isSheet ? '单元格' : '选中文字'),
              h('button', {
                className: 'dr-x',
                onClick: () => {
                  setOpen(false)
                  setEnabled(false)
                  setPending(null)
                },
                title: '关闭',
              }, '×'),
            ),
            isSheet
              ? h(
                  'div',
                  { className: 'dr-form' },
                  h(
                    'div',
                    { className: 'dr-row' },
                    h('input', {
                      value: cell.sheet,
                      placeholder: '工作表',
                      onChange: e => setCell({ ...cell, sheet: e.target.value }),
                    }),
                    h('input', {
                      value: cell.addr,
                      placeholder: '单元格，如 B3:F8',
                      onChange: e => setCell({ ...cell, addr: e.target.value }),
                    }),
                  ),
                  h('input', {
                    value: cell.value,
                    placeholder: '当前值（可选，可从预览复制）',
                    onChange: e => setCell({ ...cell, value: e.target.value }),
                  }),
                  h('textarea', {
                    value: cell.note,
                    placeholder: '修改意见',
                    onChange: e => setCell({ ...cell, note: e.target.value }),
                  }),
                  h('button', { onClick: addCell }, '加入批注'),
                )
              : h('p', { className: 'dr-hint' }, enabled
                  ? '在预览中选中文字，点击浮出的“添加批注”。可连续添加多条。'
                  : '点击上方“批注”，在预览中选中文字。'),
            h('p', { className: 'dr-tip' }, '每批建议开新会话；同一会话的历史会让每一步重读更贵。'),
            h(
              'ul',
              { className: 'dr-list' },
              comments.map(item =>
                h(
                  'li',
                  { key: item.id, className: 'dr-item' },
                  h('div', { className: 'dr-meta' }, item.cell ? `${item.sheet}!${item.cell}` : `第 ${item.page || '?'} 页`),
                  item.quote ? h('blockquote', { className: 'dr-quote' }, item.quote) : null,
                  !item.quote && item.value ? h('blockquote', { className: 'dr-quote' }, item.value) : null,
                  h('p', { className: 'dr-note' }, item.note),
                  h('button', {
                    className: 'dr-del',
                    onClick: () => commit(comments.filter(entry => entry.id !== item.id)),
                  }, '删除'),
                ),
              ),
            ),
            error ? h('p', { className: 'dr-error' }, error) : null,
            status ? h('p', { className: 'dr-status' }, status) : null,
            h(
              'footer',
              { className: 'dr-foot' },
              h('button', {
                className: 'dr-primary',
                disabled: !count || busy || !props.send,
                onClick: submit,
              }, busy ? '提交中…' : `提交 ${count} 条`),
            ),
          )
        : null

      const bubble = pending && !pending.editing
        ? h(
            'button',
            {
              className: 'dr-bubble',
              style: {
                left: Math.max(8, Math.min(pending.x, window.innerWidth - 120)),
                top: Math.max(8, Math.min(pending.y + 6, window.innerHeight - 40)),
              },
              onMouseDown: event => event.preventDefault(),
              onClick: () => setPending({ ...pending, editing: true }),
            },
            '添加批注',
          )
        : null

      const popover = pending && pending.editing
        ? h(
            'div',
            {
              className: 'dr-pop',
              'data-doc-review': 'popover',
              style: {
                left: Math.max(8, Math.min(pending.x, window.innerWidth - 320)),
                top: Math.max(8, Math.min(pending.y + 8, window.innerHeight - 240)),
              },
            },
            h('div', { className: 'dr-meta' }, pending.page ? `第 ${pending.page} 页` : '选中内容'),
            h('blockquote', { className: 'dr-quote' }, pending.quote),
            h('textarea', {
              autoFocus: true,
              value: draft,
              placeholder: '修改意见，比如“改为 2026 年 12 月底前”',
              onChange: e => setDraft(e.target.value),
            }),
            h(
              'div',
              { className: 'dr-actions' },
              h('button', {
                onClick: () => {
                  setPending(null)
                  setDraft('')
                },
              }, '取消'),
              h('button', {
                className: 'dr-primary',
                disabled: !draft.trim(),
                onClick: saveText,
              }, '保存'),
            ),
          )
        : null

      return h(React.Fragment, null, button, panel, bubble, popover)
    }

    function apply(ctx) {
      ensureStyle()
      ctx.slots.inject('sidebar.right.tab.document.actions', () =>
        ctx.slots.register(
          {
            name: 'sidebar.right.tab.document.actions',
            id: 'doc-review',
            order: -5,
            inject: sessionId => ({
              send: text => sendPrompt(ctx, sessionId, text),
            }),
          },
          ReviewAction,
        ),
      )
    }

    exports.apply = apply
    exports.inject = ['slots', 'sessions']
    return module.exports
  },
})
