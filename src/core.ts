import type { FileChild, ParagraphChild } from "docx"
import type { IChildren, IChineseFontSize, IData, IDocumentProperties, IFontSize, ISection } from "./types"
import { Document, Footer, Header, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx"

/**
 * 文档构建器
 * @author cshaptx4869
 * @date 2025-10-09
 * @link https://docx.js.org/
 */
export class DocxBuilder {
  private static count = 0
  private id = ""
  private sections: ISection[] = []
  private properties: IDocumentProperties = {}
  private defaultFont = "Arial"
  private defaultFontSize: IFontSize = 10

  constructor() {
    this.id = `docx-builder${DocxBuilder.count}`
    DocxBuilder.count++
  }

  /**
   * 设置文档属性
   * @param properties 文档属性配置
   * @returns DocxBuilder实例
   */
  setProperties(properties: IDocumentProperties) {
    this.properties = properties
    return this
  }

  /**
   * 设置默认字体
   * @param font 字体名称
   * @returns DocxBuilder实例
   */
  setDefaultFont(font: string) {
    this.defaultFont = font
    return this
  }

  /**
   * 设置默认字体大小
   * @param fontSize 字体大小，支持字符串（如"初号"）或数值（单位为磅）
   * @returns DocxBuilder实例
   */
  setDefaultFontSize(fontSize: IFontSize) {
    this.defaultFontSize = fontSize
    return this
  }

  /**
   * 添加页面
   * @param section 页面配置
   * @returns DocxBuilder实例
   */
  addSection(section: ISection) {
    this.sections.push(section)
    return this
  }

  /**
   * 添加多个页面
   * @param sections 页面配置数组
   * @returns DocxBuilder实例
   */
  addSections(sections: ISection[]) {
    this.sections.push(...sections)
    return this
  }

  /**
   * 渲染文档模板
   * @returns 渲染后的文档
   */
  render(data?: IData) {
    // 初始化文档
    const doc = new Document({
      creator: "DocxBuilder",
      ...this.properties,
      sections: this.sections.map((section) => {
        return {
          // 页眉
          headers: this.renderHeaders(section.headers, data),
          // 页脚
          footers: this.renderFooters(section.footers, data),
          // 页面属性
          properties: section.properties,
          // 页面内容
          children: this.renderChildren(section.children, data),
        }
      }),
    })

    return doc
  }

  /**
   * 渲染页眉
   * @param headers 页眉配置
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的页眉
   */
  private renderHeaders(headers?: ISection["headers"], data?: IData) {
    if (!headers) {
      return undefined
    }
    const { default: defaultHeader, first: firstHeader, even: evenHeader } = headers
    return {
      default: defaultHeader
        ? new Header({
            children: this.renderChildren(defaultHeader.children, data),
          })
        : undefined,
      first: firstHeader
        ? new Header({
            children: this.renderChildren(firstHeader.children, data),
          })
        : undefined,
      even: evenHeader
        ? new Header({
            children: this.renderChildren(evenHeader.children, data),
          })
        : undefined,
    }
  }

  /**
   * 渲染页脚
   * @param footers 页脚配置
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的页脚
   */
  private renderFooters(footers?: ISection["footers"], data?: IData) {
    if (!footers) {
      return undefined
    }
    const { default: defaultFooter, first: firstFooter, even: evenFooter } = footers
    return {
      default: defaultFooter
        ? new Footer({
            children: this.renderChildren(defaultFooter.children, data),
          })
        : undefined,
      first: firstFooter
        ? new Footer({
            children: this.renderChildren(firstFooter.children, data),
          })
        : undefined,
      even: evenFooter
        ? new Footer({
            children: this.renderChildren(evenFooter.children, data),
          })
        : undefined,
    }
  }

  /**
   * 渲染子元素
   * @param children 子元素数组
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的子元素数组
   */
  private renderChildren(children: IChildren, data?: IData) {
    const fileChild: FileChild[] = []
    children.forEach((contentItem) => {
      const { type, options } = contentItem
      if (type === "paragraph") { // 段落类型
        const paragraphChildren: ParagraphChild[] = []
        options.children.forEach((optionItem) => {
          const { type, options } = optionItem
          if (type === "text") { // 文本类型
            let value = options.text ?? "" // 默认值，文本值
            const field = options.field ?? options.htmlConfig?.field // 字段，用于从数据对象中获取值来替换默认值
            // 赋值
            if (data && field) {
              if (data[field] !== undefined) {
                // NOTE: 非 string 类型的值会被忽略
                value = String(data[field])
              }
              else if (field.includes(".")) {
                // 处理嵌套字段，如：a.b.c
                const fieldParts = field.split(".")
                let tmpValue = JSON.parse(JSON.stringify(data))
                let isBreak = false
                for (let i = 0; i < fieldParts.length; i++) {
                  if (!tmpValue[fieldParts[i]]) {
                    isBreak = true
                    break
                  }
                  tmpValue = tmpValue[fieldParts[i]]
                }
                if (!isBreak) {
                  value = String(tmpValue)
                }
              }
            }
            // 处理文本
            value !== "" && this.parseText(value).forEach((line, lineIndex) => {
              line.forEach((item, itemIndex) => {
                const text = new TextRun({
                  ...options,
                  text: item.value,
                  font: options.font ?? this.defaultFont,
                  size: options.size ?? DocxBuilder.textSize(this.defaultFontSize),
                  bold: options.bold ?? (item.type === "strong" || item.type === "b"),
                  italics: options.italics ?? (item.type === "em" || item.type === "i"),
                  underline: options.underline ?? (item.type === "u" ? { type: "single" } : undefined),
                  strike: options.strike ?? (item.type === "del" || item.type === "s"),
                  subScript: options.subScript ?? (item.type === "sub"),
                  superScript: options.superScript ?? (item.type === "sup"),
                  break: options.break ?? (lineIndex !== 0 && itemIndex === line.length - 1 ? 1 : 0),
                })
                paragraphChildren.push(text)
              })
            })
          }
          else if (type === "image") { // 图片类型
            const image = new ImageRun(options)
            paragraphChildren.push(image)
          }
        })
        // 忽略空段落
        if (paragraphChildren.length > 0) {
          const paragraph = new Paragraph({
            ...options,
            children: paragraphChildren,
          })
          fileChild.push(paragraph)
        }
      }
      else if (type === "emptyParagraph") { // 空段落类型
        const emptyParagraph = new Paragraph({
          children: [
            new TextRun({
              text: "",
            }),
          ],
        })
        fileChild.push(emptyParagraph)
      }
      else if (type === "table") { // 表格类型
        const table = new Table({
          ...options,
          // NOTE: Tables contain a list of Rows
          rows: options.rows.map((rowOptions) => {
            return new TableRow({
              ...rowOptions,
              // NOTE: Rows contain a list of TableCells
              children: rowOptions.children.map((cellOptions) => {
                return new TableCell({
                  ...cellOptions,
                  // NOTE: TableCells contain a list of Paragraphs and/or Tables. You can add Tables as tables can be nested inside each other
                  children: this.renderChildren(cellOptions.children, data),
                })
              }),
            })
          }),
        })
        fileChild.push(table)
      }
    })
    return fileChild
  }

  /**
   * 解析文本，将包含 <sup> 和 <sub> 标签的文本解析为数组
   * @param input 包含 <sup> 和 <sub> 标签的文本
   * @returns 解析后的数组，每个元素为一个对象，包含 type（类型，"sup" 或 "sub"）和 value（值）属性
   */
  private parseText(input: string) {
    const result = []
    const lines = input.split("\n")

    // 只需维护这一行：支持的内联标签
    const inlineTags = ["sup", "sub", "strong", "b", "em", "i", "del", "s", "u"]

    for (const line of lines) {
      const matches = []

      // 遍历每个标签，动态生成正则并匹配
      for (const tag of inlineTags) {
        const regex = new RegExp(`<${tag}>(.*?)</${tag}>`, "g")
        for (const match of line.matchAll(regex)) {
          matches.push({
            type: tag,
            value: match[1],
            start: match.index,
            end: match.index + match[0].length,
          })
        }
      }

      // 按出现位置排序
      matches.sort((a, b) => a.start - b.start)

      const parts = []
      let lastIndex = 0

      // 插入文本与标签片段
      for (const m of matches) {
        if (m.start > lastIndex) {
          parts.push({
            type: "text",
            value: line.slice(lastIndex, m.start),
          })
        }
        parts.push({
          type: m.type,
          value: m.value,
        })
        lastIndex = m.end
      }

      // 添加末尾剩余文本
      if (lastIndex < line.length) {
        parts.push({
          type: "text",
          value: line.slice(lastIndex),
        })
      }

      result.push(parts)
    }

    return result
  }

  /**
   * 保存文档
   * @param doc 文档对象
   * @param filename 文件名
   */
  fileSave(doc: Document, filename: string) {
    Packer.toBlob(doc).then((blob) => {
      const a = document.createElement("a")
      a.href = URL.createObjectURL(blob)
      a.download = filename.endsWith(".docx") ? filename : `${filename}.docx`
      a.click()
      URL.revokeObjectURL(a.href)
    })
  }

  /**
   * 渲染HTML模板
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的HTML字符串
   */
  renderHtml(data?: IData) {
    let html = ""
    this.sections.forEach((section, index) => {
      html += `<div class="docx-builder-section" style="page-break-after: ${index === this.sections.length - 1 ? "auto" : "always"};">`
      // 页眉
      html += `<div class="docx-builder-header">${this.renderHeadersHtml(section.headers, data)}</div>`
      // 页面内容
      html += `<div class="docx-builder-body">${this.renderChildrenHtml(section.children, data)}</div>`
      // 页脚
      html += `<div class="docx-builder-footer">${this.renderFootersHtml(section.footers, data)}</div>`
      html += `</div>`
    })
    this.injectStyle(`docx-builder-style`)
    return `<div id="${this.id}">${html}</div>`
  }

  /**
   * 注入样式
   * @param id 样式ID
   * @returns 样式元素
   */
  private injectStyle(id: string) {
    let style = document.getElementById(id)
    if (!style) {
      style = document.createElement("style")
      style.id = id
      style.textContent = `
        .docx-builder-section {
          background-color: #fff;
        }
        .docx-builder-p {
          margin: 0;
        }
        .docx-builder-table {
          border-collapse: collapse;
          width: 100%;
        }
        .docx-builder-td {
          border: 1px solid #dddddd;
          padding: 0 6px;
        }
        .docx-builder-input,
        .docx-builder-textarea,
        .docx-builder-select {
          height: 32px;
          line-height: 1.3;
          border: 1px solid #eee;
          background-color: #fff;
          color: rgba(0, 0, 0, .85);
          border-radius: 2px;
          outline: 0;
          transition: all .3s;
          box-sizing: border-box;
        }
        .docx-builder-input:hover,
        .docx-builder-textarea:hover,
        .docx-builder-select:hover {
          border-color: #d2d2d2;
        }
        .docx-builder-input:focus,
        .docx-builder-textarea:focus,
        .docx-builder-select:focus {
          border-color: #409eff;
          box-shadow: 0 0 0 3px rgba(22,183,119,.08);
        }
        .docx-builder-input {
          padding-left: 10px;
        }
        .docx-builder-textarea {
          min-height: 80px;
          line-height: 20px;
          padding: 6px 10px;
          resize: vertical;
        }
        .docx-builder-select {
          padding: 0 15px;
          cursor: pointer;
        }
        .docx-builder-option {
        }
      `
      document.head.appendChild(style)
    }
    return style
  }

  /**
   * 渲染段落HTML
   * @param children 段落内容数组
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的HTML字符串
   */
  private renderChildrenHtml(children: IChildren, data?: IData) {
    let childrenHtml = ""
    children.forEach((contentItem) => {
      const { type, options } = contentItem
      if (type === "paragraph") { // 段落类型
        let pHtml = ""
        let pStyle = `style="`
        if (options.spacing) {
          if (options.spacing.before) {
            pStyle += `padding-top: ${DocxBuilder.twipsToPx(options.spacing.before)}px;`
          }
          if (options.spacing.after) {
            pStyle += `padding-bottom: ${DocxBuilder.twipsToPx(options.spacing.after)}px;`
          }
          if (options.spacing.line) {
            pStyle += `line-height: ${options.spacing.line / 240};`
          }
        }
        pStyle += `text-align:${options.alignment ?? "left"};"`
        // 内部元素
        options.children.forEach((inlineItem) => {
          const { type, options } = inlineItem
          if (type === "text") {
            // 文本类型
            const spanStyle = `style="font-size: ${options.size ?? DocxBuilder.textSize(this.defaultFontSize)}px; font-family: ${options.font ?? this.defaultFont}; color: ${options.color ? `#${options.color}` : "#000000"};"`
            const breakTag = options.break ? "<br />".repeat(options.break) : ""
            if (options.htmlConfig) {
              // html 配置渲染
              const { field, name, props } = options.htmlConfig
              const propsStr = props ? Object.entries(props).map(([key, value]) => typeof value !== "object" ? `${key}="${value}"` : "").join(" ") : ""
              let value = options.text ?? "" // 默认值，文本值
              // 赋值
              if (data) {
                if (data[field] !== undefined) {
                  value = String(data[field])
                }
                else if (field.includes(".")) {
                  // 处理嵌套宇段，如：a.b.c
                  const fieldParts = field.split(".")
                  let tmpValue = JSON.parse(JSON.stringify(data))
                  let isBreak = false
                  for (let i = 0; i < fieldParts.length; i++) {
                    if (tmpValue[fieldParts[i]] === undefined) {
                      isBreak = true
                      break
                    }
                    tmpValue = tmpValue[fieldParts[i]]
                  }
                  if (!isBreak) {
                    value = String(tmpValue)
                  }
                }
              }
              if (name === "input") {
                // 输入框
                pHtml += `${breakTag}<input class="docx-builder-input" ${propsStr} name="${field}" value="${value}" />`
              }
              else if (name === "textarea") {
                // 文本域
                pHtml += `${breakTag}<textarea class="docx-builder-textarea" ${propsStr} name="${field}">${value}</textarea>`
              }
              else if (name === "select") {
                // 下拉选择框
                let optionHtml = ""
                if (props?.options) {
                  optionHtml += props.options.map(item => `<option class="docx-builder-option" value="${item.value}" ${value === item.value ? "selected" : ""}>${item.label}</option>`).join("")
                }
                pHtml += `${breakTag}<select class="docx-builder-select" ${propsStr} name="${field}">${optionHtml}</select>`
              }
              else {
                // 普通文本
                pHtml += `${breakTag}<span ${spanStyle}>${value.replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;").replace(/\n/g, "<br />")}</span>`
              }
            }
            else {
              // 普通文本
              pHtml += `${breakTag}<span ${spanStyle}>${options.text ? options.text.replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;").replace(/\n/g, "<br />") : ""}</span>`
            }
          }
          else if (type === "image") {
            // 图片类型
            if (options.floating === undefined) {
              const imgAttrs = `src="${options.data}" alt="${options.altText?.name ?? ""}" width="${options.transformation.width}px" height="${options.transformation.height}px"`
              pHtml += `<img ${imgAttrs} />`
            }
            // TODO 处理浮动图片
          }
        })
        childrenHtml += `<p class="docx-builder-p" ${pStyle}>${pHtml}</p>`
      }
      else if (type === "emptyParagraph") {
        // 空段落类型
        childrenHtml += "<br />"
      }
      else if (type === "table") {
        // 表格类型
        let tableHtml = ""
        options.rows.forEach((rowItem) => {
          const trStyle = `${rowItem.cantSplit ? "style=\"page-break-inside: avoid;\"" : ""}`
          let rowHtml = `<tr ${trStyle}>`
          rowItem.children.forEach((tdItem) => {
            const tdAttrs = `colspan="${tdItem.columnSpan ?? 1}" rowspan="${tdItem.rowSpan ?? 1}" style="min-width: ${tdItem.width?.size ? (typeof tdItem.width.size === "number" ? `${tdItem.width.size}cm` : tdItem.width.size) : "auto"};"`
            rowHtml += `<td class="docx-builder-td" ${tdAttrs}>${this.renderChildrenHtml(tdItem.children, data)}</td>`
          })
          rowHtml += "</tr>"
          tableHtml += rowHtml
        })
        childrenHtml += `<table class="docx-builder-table">${tableHtml}</table>`
      }
    })
    return childrenHtml
  }

  /**
   * 渲染页眉HTML
   * @param headers 页眉配置对象
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的HTML字符串
   */
  private renderHeadersHtml(headers?: ISection["headers"], data?: IData) {
    let html = ""
    if (!headers) {
      return html
    }
    const { default: defaultHeader, first: firstHeader, even: evenHeader } = headers
    if (defaultHeader) {
      html += `${this.renderChildrenHtml(defaultHeader.children, data)}`
    }
    if (firstHeader) {
      html += `${this.renderChildrenHtml(firstHeader.children, data)}`
    }
    if (evenHeader) {
      html += `${this.renderChildrenHtml(evenHeader.children, data)}`
    }
    return html
  }

  /**
   * 渲染页脚HTML
   * @param footers 页脚配置对象
   * @param data 数据对象，用于替换模板中的变量
   * @returns 渲染后的HTML字符串
   */
  private renderFootersHtml(footers?: ISection["footers"], data?: IData) {
    let html = ""
    if (!footers) {
      return html
    }
    const { default: defaultFooter, first: firstFooter, even: evenFooter } = footers
    if (defaultFooter) {
      html += `${this.renderChildrenHtml(defaultFooter.children, data)}`
    }
    if (firstFooter) {
      html += `${this.renderChildrenHtml(firstFooter.children, data)}`
    }
    if (evenFooter) {
      html += `${this.renderChildrenHtml(evenFooter.children, data)}`
    }
    return html
  }

  /**
   * 从URL获取文件的ArrayBuffer
   * @param urlStr 文件URL
   * @returns 文件的ArrayBuffer
   */
  static async fetchUrlFile(urlStr: string) {
    const url = new URL(urlStr)

    if (typeof fetch !== "undefined") {
      // 优先使用 fetch
      const response = await fetch(urlStr)
      if (!response.ok) {
        throw new Error(`Failed to fetch ${urlStr}, status code: ${response.status}`)
      }
      return await response.arrayBuffer()
    }

    // Node.js 回退
    const http = url.protocol === "https:" ? await import("node:https") : await import("node:http")

    return new Promise<ArrayBuffer>((resolve, reject) => {
      const request = http.get(urlStr, (response) => {
        if (response.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${urlStr}, status code: ${response.statusCode}`))
          response.resume()
          return
        }

        const chunks: Uint8Array[] = []
        response.on("data", (chunk: Uint8Array) => chunks.push(chunk))
        response.on("end", async () => {
          const { Buffer } = await import("node:buffer")
          // 合并所有 chunk 成一个 Buffer
          const buffer = Buffer.concat(chunks)
          // 转换为 ArrayBuffer（安全切片）
          const arrayBuffer = buffer.buffer.slice(
            buffer.byteOffset,
            buffer.byteOffset + buffer.byteLength,
          )
          resolve(arrayBuffer)
        })
      })

      request.on("error", reject)
      request.on("timeout", () => request.destroy())
      request.setTimeout(30_000) // 30秒超时
    })
  }

  /**
   * 判断字符串是否包含中文字符或中文标点符号
   * @param str 输入字符串
   * @returns 如果字符串包含中文字符或中文标点符号，则返回true；否则返回false
   */
  static isChinese(str: string) {
    // 中文和中文标点符号正则
    const chineseWithPunctRegex = /^[\u4E00-\u9FA5\u3000-\u303F\uFF00-\uFFEF]+$/
    return chineseWithPunctRegex.test(str)
  }

  /**
   * 字体大小
   * @param fontSize - 字体大小，支持字符串（如"初号"）或数值（单位为磅）
   * @returns 字体大小值
   */
  static textSize(fontSize: IFontSize) {
    // NOTE: 字体大小，数值类型的值单位是半点（half-points），1pt = 2 半点
    const fontSizeMap: Record<IChineseFontSize, number> = {
      初号: 42,
      小初: 36,
      一号: 26,
      小一: 24,
      二号: 22,
      小二: 18,
      三号: 16,
      小三: 15,
      四号: 14,
      小四: 12,
      五号: 10.5,
      小五: 9,
      六号: 7.5,
      小六: 6.5,
      七号: 5.5,
      八号: 5,
    }
    if (typeof fontSize === "number") {
      return Math.round(fontSize * 2)
    }
    return Math.round(fontSizeMap[fontSize] * 2)
  }

  /**
   * 边框大小
   * @param points - 磅值
   * @returns size值
   */
  static paragraphBorderSize(points: number) {
    // NOTE: 宽度，单位八分之一磅 (1/8 point)
    return Math.round(points * 8)
  }

  /**
   * 将行高转换为twips值
   * 240 twips = 1 行
   * @param line 行高倍数
   * @returns 行高的twips值
   */
  static paragraphSpacingLine(line: number) {
    return Math.round(line * 240)
  }

  /**
   * 图像大小
   * @param cm - 厘米值
   * @returns px值
   */
  static imageTransformationSize(cm: number) {
    // NOTE: 图像大小单位是像素 px
    return this.cmToPx(cm)
  }

  /**
   * 浮动位置偏移
   * @param cm - 厘米值
   * @returns emus值
   */
  static imageFloatingPositionOffset(cm: number) {
    // NOTE: 图像水平位置偏移量单位为EMUs
    return this.cmToEmus(cm)
  }

  /**
   * 默认单元格边距
   * @param cm - 厘米值
   * @returns twips值
   */
  static tableCellMarginSize(cm: number) {
    // NOTE: 默认单元格边距,数值类型的值单位为缇（twips，二十分之一点）
    return this.cmToTwips(cm)
  }

  /**
   * 单元格宽度
   * @param cm - 厘米值
   * @returns twips值
   */
  static tableCellWidthSize(cm: number) {
    // NOTE: 数值类型的列宽度默认单位是 twips
    return this.cmToTwips(cm)
  }

  /**
   * 表格行高
   * @param cm - 厘米值
   * @returns twips值
   */
  static tableRowHeightValue(cm: number) {
    // NOTE: 数值类型的行高默认单位是 twips
    return this.cmToTwips(cm)
  }

  /**
   * 将厘米转换为twips
   * 1 磅 = 20 twips，1 厘米 ≈ 567 twips
   * @param cm - 厘米值
   * @returns twips值
   */
  static cmToTwips(cm: number) {
    return Math.round(cm * 566.93)
  }

  /**
   * 将厘米转换为Emus
   * @param cm - 厘米值
   * @returns EMUs值
   */
  static cmToEmus(cm: number) {
    return Math.round(cm * 360000)
  }

  /**
   * 将厘米转换为像素 (基于 96 DPI)
   * @param cm - 厘米值
   * @returns  像素值
   */
  static cmToPx(cm: number) {
    return Math.round(cm * 96 / 2.54)
  }

  /**
   * 将twips转换为像素
   * @param twips - twips值
   * @returns 像素值
   */
  static twipsToPx(twips: number) {
    return Math.round(twips * (96 / 1440))
  }
}
