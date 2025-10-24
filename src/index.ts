import type {
  FileChild,
  IImageOptions,
  IParagraphOptions,
  IPropertiesOptions,
  IRunOptions,
  ISectionPropertiesOptions,
  ITableCellOptions,
  ITableOptions,
  ITableRowOptions,
  ParagraphChild,
} from "docx"
import { Document, Footer, Header, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx"

/**
 * 文档构建器
 * @author cshaptx4869
 * @date 2025-10-09
 * @link https://docx.js.org/
 */
export class DocxBuilder {
  private sections: ISection[] = []
  private properties: IDocumentProperties = {}
  private defaultFont = "Arial"
  private defaultFontSize: IFontSize = 10

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
            const field = options.htmlConfig?.field // 字段，用于从数据对象中获取值来替换默认值
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
            // 处理换行符
            value.split("\n").forEach((line, lineIndex) => {
              const text = new TextRun({
                ...options,
                text: line,
                font: options.font || this.defaultFont,
                size: options.size || DocxBuilder.textSize(this.defaultFontSize),
                break: options.break || (lineIndex === 0 ? 0 : 1),
              })
              paragraphChildren.push(text)
            })
          }
          else if (type === "image") { // 图片类型
            const image = new ImageRun(options)
            paragraphChildren.push(image)
          }
        })
        const paragraph = new Paragraph({
          ...options,
          children: paragraphChildren,
        })
        fileChild.push(paragraph)
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
      // 页眉
      html += `<div id="header-${index}">${this.renderHeadersHtml(section.headers, data)}</div>`
      // 页面内容
      html += `<div id="section-${index}" style="position: relative; page-break-after: ${index === this.sections.length - 1 ? "auto" : "always"};">${this.renderChildrenHtml(section.children, data)}</div>`
      // 页脚
      html += `<div id="footer-${index}">${this.renderFootersHtml(section.footers, data)}</div>`
    })
    return `<div id="docx-builder" style="padding: 10px; background-color: #fff; overflow: auto;">${html}</div>`
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
        const pProps = `style="text-align:${options.alignment ?? "left"};"`
        // 内部元素
        options.children.forEach((inlineItem) => {
          const { type, options } = inlineItem
          if (type === "text") {
            // 文本类型
            const spanProps = `style="font-size: ${options.size ?? DocxBuilder.textSize(this.defaultFontSize)}px; font-family: ${options.font ?? this.defaultFont}; color: ${options.color ? `#${options.color}` : "#000000"};"`
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
                pHtml += `${breakTag}<input ${propsStr} name="${field}" value="${value}" />`
              }
              else if (name === "textarea") {
                // 文本域
                pHtml += `${breakTag}<textarea ${propsStr} name="${field}">${value}</textarea>`
              }
              else if (name === "select") {
                // 下拉选择框
                let optionHtml = ""
                if (props?.options) {
                  optionHtml += props.options.map(item => `<option value="${item.value}" ${value === item.value ? "selected" : ""}>${item.label}</option>`).join("")
                }
                pHtml += `${breakTag}<select ${propsStr} name="${field}">${optionHtml}</select>`
              }
              else {
                // 普通文本
                pHtml += `${breakTag}<span ${spanProps}>${value.replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;").replace(/\n/g, "<br />")}</span>`
              }
            }
            else {
              // 普通文本
              pHtml += `${breakTag}<span ${spanProps}>${options.text ? options.text.replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;").replace(/\n/g, "<br />") : ""}</span>`
            }
          }
          else if (type === "image") {
            // 图片类型
            if (options.floating === undefined) {
              const imgProps = `src="${options.data}" alt="${options.altText?.name ?? ""}" width="${options.transformation.width}px" height="${options.transformation.height}px"`
              pHtml += `<img ${imgProps} />`
            }
            // TODO 处理浮动图片
          }
        })
        childrenHtml += `<p ${pProps}>${pHtml}</p>`
      }
      else if (type === "emptyParagraph") {
        // 空段落类型
        childrenHtml += "<br />"
      }
      else if (type === "table") {
        // 表格类型
        let tableHtml = ""
        const tableProps = `style="border-collapse: collapse; width: 100%;"`
        options.rows.forEach((rowItem) => {
          const trProps = `${rowItem.cantSplit ? "style=\"page-break-inside: avoid;\"" : ""}`
          let rowHtml = `<tr ${trProps}>`
          rowItem.children.forEach((tdItem) => {
            const tdProps = `colspan="${tdItem.columnSpan ?? 1}" rowspan="${tdItem.rowSpan ?? 1}" style="border: 1px solid #dddddd; padding: 0 6px; min-width: ${tdItem.width?.size ? (typeof tdItem.width.size === "number" ? `${tdItem.width.size}cm` : tdItem.width.size) : "auto"};"`
            rowHtml += `<td ${tdProps}>${this.renderChildrenHtml(tdItem.children, data)}</td>`
          })
          rowHtml += "</tr>"
          tableHtml += rowHtml
        })
        childrenHtml += `<table ${tableProps}>${tableHtml}</table>`
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
      return undefined
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
   * @param url 文件URL
   * @returns 文件的ArrayBuffer
   */
  static async fetchUrlFile(url: string) {
    try {
      const response = await fetch(url)
      if (!response.ok) {
        throw new Error(`Failed to fetch ${url}, status code: ${response.status}`)
      }
      return response.arrayBuffer()
    }
    catch (error: any) {
      throw new Error(`Error fetching ${url}: ${error.message}`)
    }
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
}

export default DocxBuilder

// #region 类型定义
// 段落文本元素
interface ITextType {
  type: "text"
  options: ITextTypeOptions
}
type ITextTypeOptions = IRunOptions & {
  htmlConfig?: IHtmlConfig
}
interface IHtmlConfig {
  field: string
  name?: "input" | "textarea" | "select" | "span"
  props?: {
    options?: { label: string, value: string }[]
    [key: string]: any
  }
}

// 段落图像元素
interface IImageType {
  type: "image"
  options: IImageTypeOptions
}
type IImageTypeOptions = IImageOptions

// 段落类型
export interface IParagraphType {
  type: "paragraph"
  options: IParagraphTypeOptions
}
export type IParagraphTypeOptions = Omit<IParagraphOptions, "children"> & {
  children: IParagraph
}
export type IParagraph = Array<ITextType | IImageType>

// 空段落类型
export interface IEmptyParagraphType {
  type: "emptyParagraph"
  options?: undefined
}

// 表格类型
export interface ITableType {
  type: "table"
  options: ITableTypeOptions
}
export type ITableTypeOptions = Omit<ITableOptions, "rows"> & {
  rows: ITableRows
}
export type ITableRows = ITableRow[]
export type ITableRow = Omit<ITableRowOptions, "children"> & {
  children: ITableCell[]
}
export type ITableCell = Omit<ITableCellOptions, "children"> & {
  children: IParagraphType[]
}

// 页面
export interface ISection {
  children: IChildren
  headers?: {
    default?: {
      children: IHeaderFooter
    }
    first?: {
      children: IHeaderFooter
    }
    even?: {
      children: IHeaderFooter
    }
  }
  footers?: {
    default?: {
      children: IHeaderFooter
    }
    first?: {
      children: IHeaderFooter
    }
    even?: {
      children: IHeaderFooter
    }
  }
  properties?: ISectionPropertiesOptions
}
type IChildren = Array<IParagraphType | IEmptyParagraphType | ITableType>
type IHeaderFooter = Array<IParagraphType | ITableType>
type IDocumentProperties = Omit<IPropertiesOptions, "sections">
// 数据
export type IData = Record<string, string>
// 中文字号
type IChineseFontSize = | "初号"
  | "小初"
  | "一号"
  | "小一"
  | "二号"
  | "小二"
  | "三号"
  | "小三"
  | "四号"
  | "小四"
  | "五号"
  | "小五"
  | "六号"
  | "小六"
  | "七号"
  | "八号"
type IFontSize = IChineseFontSize | number
// #endregion
