import JSZip from "jszip"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { Element as XElement, Document as XDocument } from "@xmldom/xmldom"
import fs from "fs"
import path from "path"

const TEMPLATE_PATH = path.join(process.cwd(), "public", "formato-examen.xlsx")

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

const EXAM_CHECKBOX_MAP: Record<string, number> = {
  "EXAMEN MEDICO OCUPACIONAL OSTEOMUSCULAR": 7,
  "CUADRO HEMATICO": 17,
  "GLICEMIA EN AYUNAS": 16,
  "PERFIL LIPIDICO COMPLETO": 15,
  "TRANSAMINASAS (ALAT/ASAT)": 14,
  MICROALBUMINURIA: 13,
  AUDIOMETRIA: 12,
  OPTOMETRIA: 8,
  "PARCIAL DE ORINA": 23,
  "CAMPIMETRIA POR CONFRONTACION": 24,
  "PRUEBA DE DROGAS EN ORINA (PANEL X 5 DROGAS)": 27,
  ESPIROMETRIA: 18,
  "RX DE TORAX AP Y LATERAL": 19,
  "ELECTROCARDIOGRAMA (MAYORES DE 45)": 20,
  "PRUEBA NEUROPSICOLOGICA PARA ALTURAS": 9,
  "PRUEBA PSICOSENSOMETRICA": 10,
  "PRUEBAS PSICOSENSOMETRICAS": 11,
  COPROLOGICO: 21,
  "PRUEBA DE EMBARAZO": 22,
  "VACUNA TETANO": 25,
  "VACUNA FIEBRE AMARILLA": 26,
  OTRO: 28,
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
  nombre_empresa: { cell: "D7", style: "96" },
  nit: { cell: "I7", style: "93" },
  empresa_mision: { cell: "D8", style: "85" },
  centro_costos: { cell: "D9", style: "88" },
  ciudad_proyecto: { cell: "I9", style: "90" },
  nombre_aspirante: { cell: "D13", style: "58" },
  cedula: { cell: "D14", style: "59" },
  cargo: { cell: "D15", style: "67" },
  nombre_ips: { cell: "D20", style: "81" },
  ciudad_ips: { cell: "D21", style: "72" },
  telefono_ips: { cell: "I21", style: "75" },
  direccion_ips: { cell: "D22", style: "75" },
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

export async function generateExcel(data: FormData): Promise<Buffer> {
  const templateBuffer = fs.readFileSync(TEMPLATE_PATH)
  const zip = await JSZip.loadAsync(templateBuffer)

  const sheetXmlStr = await zip.file("xl/worksheets/sheet1.xml")!.async("string")
  const parser = new DOMParser()
  const doc = parser.parseFromString(sheetXmlStr, "text/xml")
  const root = doc.documentElement
  if (!root) throw new Error("Failed to parse sheet XML")
  const serializer = new XMLSerializer()

  const sheetData = findChildNS(root, "sheetData")
  if (!sheetData) throw new Error("sheetData not found")

  for (const [key, mapping] of Object.entries(CELL_MAP)) {
    const rawValue = (data as unknown as Record<string, string>)[key]
    const value = rawValue ?? ""

    const { col, colNum, row: rowNum } = parseCellRef(mapping.cell)
    let finalValue: string = value

    if (key === "cedula") {
      finalValue = value.replace(/\./g, "")
    }

    let rowEl = findRowByNum(sheetData, rowNum)
    if (!rowEl) {
      rowEl = doc.createElementNS(NS, "row") as XElement
      rowEl.setAttribute("r", String(rowNum))
      insertRowSorted(sheetData, rowEl, rowNum)
    }

    removeCellByRef(rowEl, mapping.cell)

    const cellEl = doc.createElementNS(NS, "c") as XElement
    cellEl.setAttribute("r", mapping.cell)
    cellEl.setAttribute("s", mapping.style)

    if (key === "hora_examen") {
      if (value.trim()) {
        const decimal = formatHoraDecimal(value)
        const vEl = doc.createElementNS(NS, "v") as XElement
        vEl.textContent = String(decimal)
        cellEl.appendChild(vEl)
      }
    } else {
      cellEl.setAttribute("t", "inlineStr")
      const isEl = doc.createElementNS(NS, "is") as XElement
      const tEl = doc.createElementNS(NS, "t") as XElement
      tEl.textContent = finalValue
      isEl.appendChild(tEl)
      cellEl.appendChild(isEl)
    }

    insertCellSorted(rowEl, cellEl, colNum)
  }

  // "Cual?" field
  if (data.tipo_examen_otro) {
    let rowEl = findRowByNum(sheetData, 29)
    if (!rowEl) {
      rowEl = doc.createElementNS(NS, "row") as XElement
      rowEl.setAttribute("r", "29")
      insertRowSorted(sheetData, rowEl, 29)
    }
    removeCellByRef(rowEl, "D29")
    const cellEl = doc.createElementNS(NS, "c") as XElement
    cellEl.setAttribute("r", "D29")
    cellEl.setAttribute("s", "8")
    cellEl.setAttribute("t", "inlineStr")
    const isEl = doc.createElementNS(NS, "is")
    const tEl = doc.createElementNS(NS, "t")
    tEl.textContent = data.tipo_examen_otro
    isEl.appendChild(tEl)
    cellEl.appendChild(isEl)
    insertCellSorted(rowEl, cellEl, colLetterToNum("D"))
  }

  // Aptitudes médicas especiales (X marks)
  if (data.aptitudes) {
    const aptMap: Record<string, string> = {
      "TRABAJO EN ALTURAS": "D50",
      "TRABAJO EN ESPACIOS CONFINADOS": "D51",
      CONDUCCIÓN: "F50",
      "TRABAJO EN CALIENTE": "F51",
    }

    const aptitudesList = data.aptitudes
      .split(",")
      .map((a) => a.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim())

    for (const [aptKey, cellRef] of Object.entries(aptMap)) {
      const normKey = aptKey.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase()
      const matched = aptitudesList.some((a) => a === normKey || a.includes(normKey) || normKey.includes(a))
      const { row: aptRow, colNum: aptColNum } = parseCellRef(cellRef)

      let aptRowEl = findRowByNum(sheetData, aptRow)
      if (!aptRowEl) {
        aptRowEl = doc.createElementNS(NS, "row") as XElement
        aptRowEl.setAttribute("r", String(aptRow))
        insertRowSorted(sheetData, aptRowEl, aptRow)
      }

      removeCellByRef(aptRowEl, cellRef)

      const aptCellEl = doc.createElementNS(NS, "c") as XElement
      aptCellEl.setAttribute("r", cellRef)
      aptCellEl.setAttribute("s", "40")
      aptCellEl.setAttribute("t", "inlineStr")
      const aptIsEl = doc.createElementNS(NS, "is") as XElement
      const aptTEl = doc.createElementNS(NS, "t") as XElement
      aptTEl.textContent = matched ? "X" : ""
      aptIsEl.appendChild(aptTEl)
      aptCellEl.appendChild(aptIsEl)
      insertCellSorted(aptRowEl, aptCellEl, aptColNum)
    }
  }

  const updatedSheetXml = serializer.serializeToString(doc)
  await updateZipFile(zip, "xl/worksheets/sheet1.xml", updatedSheetXml)

  // Checkboxes
  for (let i = 1; i <= 6; i++) {
    await setCheckbox(zip, i, false)
  }
  const tipoCbId = TIPO_EXAMEN_MAP[data.tipo_examen.toUpperCase()]
  if (tipoCbId) await setCheckbox(zip, tipoCbId, true)

  for (const cbId of Object.values(EXAM_CHECKBOX_MAP)) {
    await setCheckbox(zip, cbId, false)
  }
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

// --- XML DOM Helpers ---

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
  for (const el of toRemove) {
    rowEl.removeChild(el)
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

async function setCheckbox(
  zip: JSZip,
  cbId: number,
  checked: boolean
): Promise<void> {
  const propPath = `xl/ctrlProps/ctrlProp${cbId}.xml`
  const propFile = zip.file(propPath)
  if (!propFile) return

  let propXml = await propFile.async("string")
  propXml = propXml.replace(/\s*checked="[^"]*"/g, "")
  if (checked) {
    propXml = propXml.replace(
      /<formControlPr\s/,
      '<formControlPr checked="Checked" '
    )
  }
  await updateZipFile(zip, propPath, propXml)
}

async function updateZipFile(
  zip: JSZip,
  filePath: string,
  content: string | Buffer
): Promise<void> {
  zip.remove(filePath)
  zip.file(filePath, content)
}
