import { unzipSync, type UnzipFileInfo } from "fflate"
import { streamXML, SmolXMLNode } from "./smol-xml.ts"

const ZERO = "0".charCodeAt(0)
const A = "A".charCodeAt(0)
const a = "a".charCodeAt(0)

// For an A1 reference, return the row and column numbers in the spreadsheet
export function a1_to_row_col(a1: string) {
  let col = 0
  let row = 0

  for (const c of a1) {
    const code = c.charCodeAt(0)
    if (code < A) {
      // number
      row = row * 10 + (code - ZERO)
    } else if (code < a){
      col = (col * 26) + (code - A + 1)
    } else {
      col = (col * 26) + (code - a + 1)
    }
  }

  // A is really 1 initially in excel terms
  return { row, col }
}

/**
  Unzip files in the .xlsx and return their text content. We can filter the files by name or by content.
  @param file - The .xlsx file to unzip
  @param filter - A function that replies whether or not to read the file
  @returns A dictionary of the files and their text content
*/
function unzip(file: Uint8Array<ArrayBuffer>, filter: (file: UnzipFileInfo) => boolean): {[name: string]: string} {
  const txt = new TextDecoder()
  const files: { [key: string]: string } = {}
  const res = unzipSync(file, {
    filter: (file) => {
      return filter(file)
    }
  })
  for (const [key, value] of Object.entries(res)) {
    files[key] = txt.decode(value)
  }
  return files
}

/**
  Used to apply the tint to the hex color that comes from a theme in a .xlsx file
  @param hex - The hex color to apply the tint to
  @param tint - The tint to apply to the hex color
  @returns The hex color with the tint applied
*/
function applyExcelTint(hex: string, tint: number) {
  // Parse R, G, B
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);

  function tintChannel(channel: number) {
    if (tint < 0) {
      // Darken
      return Math.round(channel * (1 + tint));
    } else {
      // Lighten
      return Math.round(channel + (255 - channel) * tint);
    }
  }

  r = Math.min(255, Math.max(0, tintChannel(r)));
  g = Math.min(255, Math.max(0, tintChannel(g)));
  b = Math.min(255, Math.max(0, tintChannel(b)));

  // Convert back to #RRGGBB
  return (
    r.toString(16).padStart(2, '0') +
    g.toString(16).padStart(2, '0') +
    b.toString(16).padStart(2, '0')
  ).toUpperCase();
}

export interface SmolCell {
  A1: string
  row: number
  col: number
  v: string | number | boolean | null
  error?: string
}


class SmolSheet {
  constructor(
    public name: string, public visible: boolean, public id: string, public data: Map<string, SmolCell>) { }

  get(row: number, col: number): SmolCell | undefined {
    return this.data.get(`${row}:${col}`)
  }

  min_row: number = Infinity
  max_row: number = -Infinity
  min_col: number = Infinity
  max_col: number = -Infinity
}


export class SmolWorkbook {
  constructor() { }

  sheets: SmolSheet[] = []
  sheets_by_name: Map<string, SmolSheet> = new Map()

  async read(file: Uint8Array<ArrayBuffer>, keep_sheet?: (sheet: SmolSheet) => boolean) {
    // console.time("read-mine")

    const init_files = new Set([
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/sharedStrings.xml",
      "xl/styles.xml",
      "xl/theme/theme1.xml",
    ])

    const txt = new TextDecoder()
    const files: { [key: string]: string } = {}
    Object.assign(files, unzip(file, (file) => init_files.has(file.name)))

    const rels: {[rId: string]: string} = {}
    const rels_xml = files["xl/_rels/workbook.xml.rels"]
    streamXML("Relationship", rels_xml, (rel) => {
      rels[rel.attrs.Id ?? ""] = rel.attrs.Target ?? ""
    })

    const files_to_read = new Set<string>()
    const wb_str = files["xl/workbook.xml"]
    streamXML("sheet", wb_str, (st) => {
      const sht = new SmolSheet(
        st.attrs.name ?? "",
        st.attrs.state === "visible",
        st.attrs["r:id"] ?? "",
        new Map<string, SmolCell>())
      if (keep_sheet?.(sht) === false) { return }
      this.sheets.push(sht)
      files_to_read.add("xl/" + rels[sht.id])
    })

    // Now, re-read the files we need with the collections
    Object.assign(files, unzip(file,
      (file) => files_to_read.has(file.name)
    ))

    // build the shared strings
    let strs: string[] = []
    try {
      const strings = files["xl/sharedStrings.xml"]
      streamXML("si", strings, (node) => {
        strs.push(node.textContent)
      })
    } catch (e) {
    }

    const theme = files["xl/theme/theme1.xml"]
    const theme_colors = new Map<number, string>()
    if (theme) {
      streamXML("a:clrScheme", theme, node => {
        for (let i = 0, l = node.children.length; i < l; i++) {
          const c = (node.children[i] as SmolXMLNode)?.children[0] as SmolXMLNode
          if (!c) { continue }
          if (c.tag === "a:srgbClr" && c.attrs.val) {
            theme_colors.set(i, c.attrs.val ?? "")
          }
        }
      })
    }

    const style = files["xl/styles.xml"]
    // console.log("STYLE", new DOMParser().parseFromString(style, "application/xml"))
    const bg_style = new Map<string, string>()
    const fills: string[] = []

    streamXML(["cellXfs", "patternFill"], style, (node) => {
      if (node.tag === "patternFill") {
        if (node.attrs.patternType === "solid") {
          for (let c of node.children) {
            if (c instanceof SmolXMLNode && c.tag === "fgColor") {
              if (c.attrs.rgb) {
                fills.push((c.attrs.rgb ?? "").slice(2))
              } else if (c.attrs.theme) {
                const tint = Number(c.attrs.tint ?? 1)
                const color = theme_colors.get(Number(c.attrs.theme ?? 0))
                if (color) {
                  fills.push(applyExcelTint(color, tint))
                }
              } else {
                fills.push("")
              }
              return
            }
          }
        }
        fills.push("")
        return
      }
      if (node.tag === "cellXfs") {
        for (let i = 0, l = node.children.length; i < l; i++) {
          const xf = node.children[i] as SmolXMLNode
          if (xf.attrs.fillId && xf.attrs.fillId !== "0") {
            bg_style.set(""+i, fills[Number(xf.attrs.fillId)])
          }
        }
      }
      // console.log(node)
    })
    // console.log(fills)
    // console.log(bg_style)


    // console.log(style)

    for (let s of this.sheets) {
      const sheet_str = files[`xl/${rels[s.id]}`]
      streamXML("c", sheet_str, (node) => {
        const at = node.attrs
        if (!at.r) { return }

        const v = node.children.find((c): c is SmolXMLNode => c instanceof SmolXMLNode && c.tag === "v")?.textContent
        const {row, col} = a1_to_row_col(at.r)
        const c = at.s ?? ""
        const cell_obj = {v: null, row, col, A1: at.r} as SmolCell
        if (at.t === "s") {
          cell_obj.v = strs[Number(v)] ?? ""
        } else if (at.t === "b") {
          cell_obj.v = !!v
        } else if (at.t === "n") {
          cell_obj.v = Number(v)
        } else if (at.t === "e") {
          cell_obj.error = v
        } else if (!at.t) {
          if (v != null) {
            cell_obj.v = Number(v)
          }
        } else if (at.t === "inlineStr") {
          // they may need to be handled differently
          cell_obj.v = v ?? ""
        } else if (at.t === "str") {
          cell_obj.v = v ?? ""
        }

        if (cell_obj.v == null) {
          let styl = bg_style.get(at.s ?? "")
          if (styl) {
            cell_obj.v = "#" + styl
          }
        }

        // console.log(at.s)

        // Only if added to the sheet
        if (cell_obj.v != null) {
          s.data.set(`${row}:${col}`, cell_obj)
          s.min_row = Math.min(s.min_row, row)
          s.max_row = Math.max(s.max_row, row)
          s.min_col = Math.min(s.min_col, col)
          s.max_col = Math.max(s.max_col, col)
          // #FFFF00
        }
      })

      this.sheets_by_name.set(s.name, s)

    }
  }
}