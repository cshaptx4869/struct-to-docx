import type {
  IImageOptions,
  IParagraphOptions,
  IPropertiesOptions,
  IRunOptions,
  ISectionPropertiesOptions,
  ITableCellOptions,
  ITableOptions,
  ITableRowOptions,
} from "docx"

// 段落文本元素
interface ITextType {
  type: "text"
  options: ITextTypeOptions
}
type ITextTypeOptions = IRunOptions & {
  field?: string
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
export type IChildren = Array<IParagraphType | IEmptyParagraphType | ITableType>
export type IHeaderFooter = Array<IParagraphType | ITableType>

// 文档属性
export type IDocumentProperties = Omit<IPropertiesOptions, "sections">

// 中文字号
export type IChineseFontSize = | "初号"
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
export type IFontSize = IChineseFontSize | number

// 数据
export type IData = Record<string, string>
