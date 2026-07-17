import JSZip from "jszip"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { Element as XElement } from "@xmldom/xmldom"
import fs from "fs"
import path from "path"

const TEMPLATE_PATH = path.join(process.cwd(), "public", "formato-examen version 2.xlsx")

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

const EXAM_CHECKBOX_MAP: Record<string, number> = {
  "EXAMEN MEDICO OCUPACIONAL OSTEOMUSCULAR": 7,
  "CUADRO HEMATICO": 17,
  "GLICEMIA EN AYUNAS": 16,
  "PERFIL LIPIDICO COMPLETO": 15,
  "TRANSAMINASAS (ALAT/ASAT)": 14,
  AUDIOMETRIA: 12,
  OPTOMETRIA: 8,
  "PARCIAL DE ORINA": 23,
  ESPIROMETRIA: 18,
  "RX DE TORAX AP Y LATERAL": 19,
  "ELECTROCARDIOGRAMA (APLICAR MAYORES DE 50)": 20,
  "PRUEBA NEUROPSICOLOGICA PARA ALTURAS": 9,
  "PRUEBA PSICOSENSOMETRICA": 10,
  "PRUEBA DE EQUILIBRIO": 11,
  "KOH DE UNAS": 13,
  "PRUEBA DROGAS EN ORINA (PANEL X5)": 21,
  "PRUEBA DE ALCOHOL": 22,
  "VACUNA TETANO": 24,
  "VACUNA FIEBRE AMARILLA": 25,
  "ANTIGENO PROSTATICO-PSA (MAYORES 35 ANOS)": 26,
  "PRUEBA TB (TUBERCULOSIS) EN SANGRE": 27,
  "RX COLUMNA LUMBAR": 28,
  MICROALBUMINURIA: 29,
}

const TIPO_EXAMEN_MAP: Record<string, number> = {
  INGRESO: 1,
  PERIODICO: 2,
  EGRESO: 3,
  POSTINCAPACIDAD: 4,
  "CAMBIO DE LABOR U OFICIO": 5,
  OTRO: 6,
}

export interface FormData {
  proyecto: string
  tipo_examen: string
  tipo_examen_otro: string
  cargo: string
  nombre_aspirante: string
  cedula: string
  nombre_empresa: string
  nit: string
  empresa_mision: string
  centro_costos: string
  ciudad_proyecto: string
  nombre_ips: string
  ciudad_ips: string
  telefono_ips: string
  direccion_ips: string
  fecha_examen: string
  hora_examen: string
  examenes: string[]
  aptitudes: string
}

interface CellMapping {
  cell: string
  style: string
}

const CELL_MAP: Record<string, CellMapping> = {
  nombre_empresa: { cell: "D7", style: "90" },
  nit: { cell: "I7", style: "4" },
  empresa_mision: { cell: "D8", style: "85" },
  centro_costos: { cell: "D9", style: "88" },
  ciudad_proyecto: { cell: "I9", style: "90" },
  nombre_aspirante: { cell: "D13", style: "59" },
  cedula: { cell: "D14", style: "60" },
  cargo: { cell: "D15", style: "68" },
  nombre_ips: { cell: "D20", style: "79" },
  ciudad_ips: { cell: "D21", style: "73" },
  telefono_ips: { cell: "I21", style: "76" },
  direccion_ips: { cell: "D22", style: "61" },
  fecha_examen: { cell: "D23", style: "84" },
  hora_examen: { cell: "I23", style: "83" },
}

function formatHoraDecimal(hora: string): number {
  const match = hora.match(/(\d{1,2}):(\d{2})/)
  if (!match) return 0
  let hours = parseInt(match[1])
  const mins = parseInt(match[2])
  if (hora.toLowerCase().includes("p") && hours < 12) hours += 12
  if (hora.toLowerCase().includes("a") && hours === 12) hours = 0
  return (hours * 60 + mins) / (24 * 60)
}

function colLetterToNum(col: string): number {
  let result = 0
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64)
  }
  return result
}

function parseCellRef(ref: string): { col: string; colNum: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) throw new Error(`Invalid cell ref: ${ref}`)
  return { col: match[1], colNum: colLetterToNum(match[1]), row: parseInt(match[2]) }
}

// --- Shared Strings Manager ---
class SharedStrings {
  items: string[]

  constructor(xml: string) {
    this.items = []
    const siRegex = /<si>[\s\S]*?<\/si>/g
    let m
    while ((m = siRegex.exec(xml)) !== null) {
      const tMatch = m[0].match(/<t[^>]*>([^<]*)<\/t>/)
      this.items.push(tMatch ? tMatch[1] : "")
    }
  }

  getOrAdd(value: string): number {
    const idx = this.items.indexOf(value)
    if (idx >= 0) return idx
    this.items.push(value)
    return this.items.length - 1
  }

  toXml(): string {
    const count = this.items.length
    const uniqueCount = new Set(this.items).size
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
    xml += `<sst xmlns="${NS}" count="${count}" uniqueCount="${uniqueCount}">`
    for (const item of this.items) {
      xml += `<si><t>${escXml(item)}</t></si>`
    }
    xml += `</sst>`
    return xml
  }
}

function escXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

// --- Main Generator ---
export async function generateExcel(data: FormData): Promise<Buffer> {
  const templateBuffer = fs.readFileSync(TEMPLATE_PATH)
  const zip = await JSZip.loadAsync(templateBuffer)

  // 1. Load shared strings
  const ssXml = await zip.file("xl/sharedStrings.xml")!.async("string")
  const sharedStrings = new SharedStrings(ssXml)

  // 2. Build list of values to add to shared strings
  const cellValues: { key: string; ref: string; style: string; ssIdx?: number; numValue?: number }[] = []

  for (const [key, mapping] of Object.entries(CELL_MAP)) {
    const rawValue = (data as unknown as Record<string, string>)[key]
    const value = (rawValue ?? "").trim()
    if (!value) continue

    let finalValue = value
    if (key === "cedula") finalValue = value.replace(/\./g, "")

    if (key === "hora_examen") {
      cellValues.push({ key, ref: mapping.cell, style: mapping.style, numValue: formatHoraDecimal(value) })
    } else {
      const idx = sharedStrings.getOrAdd(finalValue)
      cellValues.push({ key, ref: mapping.cell, style: mapping.style, ssIdx: idx })
    }
  }

  if (data.tipo_examen_otro) {
    const idx = sharedStrings.getOrAdd(data.tipo_examen_otro)
    cellValues.push({ key: "cual", ref: "D29", style: "8", ssIdx: idx })
  }

  // Aptitudes
  const aptMap: Record<string, { cell: string; style: string }> = {
    "TRABAJO EN ALTURAS": { cell: "D51", style: "41" },
    "TRABAJO EN ESPACIOS CONFINADOS": { cell: "D52", style: "93" },
    CONDUCCIÓN: { cell: "F51", style: "41" },
    "TRABAJO EN CALIENTE": { cell: "F52", style: "94" },
  }

  if (data.aptitudes) {
    const aptitudesList = data.aptitudes
      .split(",")
      .map((a) => a.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim())
    const xIdx = sharedStrings.getOrAdd("X")

    for (const [aptKey, { cell: cellRef, style }] of Object.entries(aptMap)) {
      const normKey = aptKey.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase()
      const matched = aptitudesList.some((a) => a === normKey || a.includes(normKey) || normKey.includes(a))
      if (matched) {
        cellValues.push({ key: `apt_${aptKey}`, ref: cellRef, style, ssIdx: xIdx })
      }
    }
  }

  // 3. Update shared strings in ZIP
  await updateZipFile(zip, "xl/sharedStrings.xml", sharedStrings.toXml())

  // 4. Modify sheet1.xml
  const sheetXmlStr = await zip.file("xl/worksheets/sheet1.xml")!.async("string")
  const parser = new DOMParser()
  const doc = parser.parseFromString(sheetXmlStr, "text/xml")
  const root = doc.documentElement
  if (!root) throw new Error("Failed to parse sheet XML")
  const serializer = new XMLSerializer()

  const sheetData = findChildNS(root, "sheetData")
  if (!sheetData) throw new Error("sheetData not found")

  // Collect all cell refs that will be written (for clearing template defaults)
  const writtenRefs = new Set(cellValues.map((v) => v.ref))

  // Also clear template cells that are NOT being written (empty user fields)
  for (const [key, mapping] of Object.entries(CELL_MAP)) {
    const rawValue = (data as unknown as Record<string, string>)[key]
    const value = (rawValue ?? "").trim()
    if (value) continue // already handled above
    writtenRefs.add(mapping.cell) // will be cleared
  }

  // Clear aptitudes cells
  for (const { cell: cellRef } of Object.values(aptMap)) {
    writtenRefs.add(cellRef)
  }

  // Remove old cells and insert new ones
  for (const cv of cellValues) {
    const { colNum, row: rowNum } = parseCellRef(cv.ref)
    let rowEl = findRowByNum(sheetData, rowNum)
    if (!rowEl) {
      rowEl = doc.createElementNS(NS, "row") as XElement
      rowEl.setAttribute("r", String(rowNum))
      insertRowSorted(sheetData, rowEl, rowNum)
    }

    // Try to reuse existing cell
    const existingCell = findCellByRef(rowEl, cv.ref)

    if (existingCell) {
      // Modify existing cell in-place preserving all attributes
      if (cv.ssIdx !== undefined) {
        existingCell.setAttribute("t", "s")
        setOrReplaceChild(existingCell, "v", String(cv.ssIdx))
      } else if (cv.numValue !== undefined) {
        existingCell.removeAttribute("t")
        setOrReplaceChild(existingCell, "v", String(cv.numValue))
      }
    } else {
      // Create new cell
      const cellEl = doc.createElementNS(NS, "c") as XElement
      cellEl.setAttribute("r", cv.ref)
      cellEl.setAttribute("s", cv.style)

      if (cv.ssIdx !== undefined) {
        cellEl.setAttribute("t", "s")
        const vEl = doc.createElementNS(NS, "v") as XElement
        vEl.textContent = String(cv.ssIdx)
        cellEl.appendChild(vEl)
      } else if (cv.numValue !== undefined) {
        const vEl = doc.createElementNS(NS, "v") as XElement
        vEl.textContent = String(cv.numValue)
        cellEl.appendChild(vEl)
      }

      insertCellSorted(rowEl, cellEl, colNum)
    }
  }

  // Clear template cells whose values should be empty (user left them blank)
  for (const ref of writtenRefs) {
    if (!cellValues.some((cv) => cv.ref === ref)) {
      const { row: rowNum } = parseCellRef(ref)
      const rowEl = findRowByNum(sheetData, rowNum)
      if (rowEl) {
        const existingCell = findCellByRef(rowEl, ref)
        if (existingCell) {
          // Clear value but keep cell (preserves formatting)
          existingCell.removeAttribute("t")
          const toRemove: XElement[] = []
          const children = existingCell.childNodes
          for (let i = 0; i < children.length; i++) {
            const child = children[i] as unknown as XElement
            if (child.nodeType === 1) toRemove.push(child)
          }
          for (const el of toRemove) existingCell.removeChild(el)
        }
      }
    }
  }

  const updatedSheetXml = serializer.serializeToString(doc)
  await updateZipFile(zip, "xl/worksheets/sheet1.xml", updatedSheetXml)

  // 5. Checkboxes
  for (let i = 1; i <= 6; i++) await setCheckbox(zip, i, false)
  const tipoCbId = TIPO_EXAMEN_MAP[data.tipo_examen.toUpperCase()]
  if (tipoCbId) await setCheckbox(zip, tipoCbId, true)

  for (const cbId of Object.values(EXAM_CHECKBOX_MAP)) await setCheckbox(zip, cbId, false)
  const examSet = new Set(data.examenes.map((e) => e.trim().toUpperCase()))
  for (const [examName, cbId] of Object.entries(EXAM_CHECKBOX_MAP)) {
    if (examSet.has(examName)) await setCheckbox(zip, cbId, true)
  }

  const result = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 5 },
  })

  return result as Buffer
}

// --- DOM Helpers ---

function findChildNS(parent: XElement, localName: string): XElement | null {
  const children = parent.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === localName || child.nodeName === localName)) {
      return child
    }
  }
  return null
}

function findRowByNum(sheetData: XElement, rowNum: number): XElement | null {
  const children = sheetData.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "row" || child.nodeName === "row")) {
      const r = parseInt(child.getAttribute("r") || "0")
      if (r === rowNum) return child
    }
  }
  return null
}

function insertRowSorted(sheetData: XElement, newRow: XElement, rowNum: number): void {
  const children = sheetData.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "row" || child.nodeName === "row")) {
      const r = parseInt(child.getAttribute("r") || "0")
      if (rowNum < r) {
        sheetData.insertBefore(newRow, child)
        return
      }
    }
  }
  sheetData.appendChild(newRow)
}

function removeCellByRef(rowEl: XElement, ref: string): void {
  const toRemove: XElement[] = []
  const children = rowEl.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "c" || child.nodeName === "c")) {
      if (child.getAttribute("r") === ref) {
        toRemove.push(child)
      }
    }
  }
  for (const el of toRemove) rowEl.removeChild(el)
}

function findCellByRef(rowEl: XElement, ref: string): XElement | null {
  const children = rowEl.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "c" || child.nodeName === "c")) {
      if (child.getAttribute("r") === ref) return child
    }
  }
  return null
}

function setOrReplaceChild(parent: XElement, tagName: string, textContent: string): void {
  // Remove existing child with this tagName
  const toRemove: XElement[] = []
  const children = parent.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === tagName || child.nodeName === tagName)) {
      toRemove.push(child)
    }
  }
  for (const el of toRemove) parent.removeChild(el)

  // Also remove <is> child (for inline strings)
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "is" || child.nodeName === "is")) {
      parent.removeChild(child)
    }
  }

  // Create new child
  const newEl = parent.ownerDocument?.createElementNS(NS, tagName) as XElement
  if (newEl) {
    newEl.textContent = textContent
    parent.appendChild(newEl)
  }
}

function insertCellSorted(rowEl: XElement, newCell: XElement, colNum: number): void {
  const children = rowEl.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as unknown as XElement
    if (child.nodeType === 1 && (child.localName === "c" || child.nodeName === "c")) {
      const ref = child.getAttribute("r") || ""
      const { colNum: existingColNum } = parseCellRef(ref)
      if (colNum < existingColNum) {
        rowEl.insertBefore(newCell, child)
        return
      }
    }
  }
  rowEl.appendChild(newCell)
}

// --- Checkbox Helpers ---

async function setCheckbox(zip: JSZip, cbId: number, checked: boolean): Promise<void> {
  const propPath = `xl/ctrlProps/ctrlProp${cbId}.xml`
  const propFile = zip.file(propPath)
  if (!propFile) return
  let propXml = await propFile.async("string")
  propXml = propXml.replace(/\s*checked="[^"]*"/g, "")
  if (checked) {
    propXml = propXml.replace(/<formControlPr\s/, '<formControlPr checked="Checked" ')
  }
  await updateZipFile(zip, propPath, propXml)
}

async function updateZipFile(zip: JSZip, filePath: string, content: string | Buffer): Promise<void> {
  zip.remove(filePath)
  zip.file(filePath, content)
}
