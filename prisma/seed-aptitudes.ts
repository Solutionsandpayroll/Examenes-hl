import "dotenv/config"
import { PrismaClient } from "../src/generated/prisma/client"
import { PrismaPg } from "@prisma/adapter-pg"
import ExcelJS from "exceljs"
import path from "path"

const prisma = new PrismaClient({
  adapter: new PrismaPg({ connectionString: process.env.DATABASE_URL! }),
})

async function readExcelRows(filePath: string): Promise<Record<string, string>[]> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)
  const worksheet = workbook.worksheets[0]
  const rows: Record<string, string>[] = []

  const headers: string[] = []
  worksheet.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber] = String(cell.value ?? "").trim()
  })

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    const obj: Record<string, string> = {}
    row.eachCell((cell, colNumber) => {
      obj[headers[colNumber]] = String(cell.value ?? "").trim()
    })
    if (Object.values(obj).some((v) => v)) rows.push(obj)
  })

  return rows
}

async function main() {
  const dataDir = path.join(process.cwd(), "prisma", "data")
  const entradaPath = path.join(dataDir, "examenes-medicos-entrada.xlsx")
  const periodicosPath = path.join(dataDir, "examenes-medicos-periodicos.xlsx")

  console.log("Reading Excel files...")
  const entradaRows = await readExcelRows(entradaPath)
  const periodicosRows = await readExcelRows(periodicosPath)

  // Collect aptitudes per cargo+proyecto
  const aptMapIngreso = new Map<string, string>()
  const aptMapPeriodico = new Map<string, string>()

  for (const row of entradaRows) {
    const proyectoSiga = row["Proyecto"] || row["PROYECTO"]
    const cargoSigla = row["Título"] || row["TITULO"] || row["Title"] || row["TITLE"]
    if (!proyectoSiga || !cargoSigla) continue

    const apt = (row["Aptitudes"] || row["APTITUDES"] || "").trim()
    if (apt) {
      const key = `${proyectoSiga.trim()}|${cargoSigla.trim()}`
      aptMapIngreso.set(key, apt)
    }
  }

  for (const row of periodicosRows) {
    const proyectoPr = row["Proyecto_pr"] || row["PROYECTO_PR"]
    const proyectoSiga = row["Proyecto"] || row["PROYECTO"] || proyectoPr
    const cargoSigla = row["Title"] || row["TITLE"] || row["Título"] || row["TITULO"]
    if (!proyectoSiga || !cargoSigla) continue

    const apt = (row["Aptitudes"] || row["APTITUDES"] || "").trim()
    if (apt) {
      const key = `${proyectoSiga.trim()}|${cargoSigla.trim()}`
      aptMapPeriodico.set(key, apt)
    }
  }

  console.log(`Found ${aptMapIngreso.size} ingreso aptitudes`)
  console.log(`Found ${aptMapPeriodico.size} periodico aptitudes`)

  // Load all cargos with their proyectos
  const cargos = await prisma.cargo.findMany({
    include: { proyecto: true },
  })

  let updatedIngreso = 0
  let updatedPeriodico = 0

  for (const cargo of cargos) {
    const key = `${cargo.proyecto.nombre}|${cargo.nombre}`

    const aptI = aptMapIngreso.get(key)
    if (aptI) {
      await prisma.cargo.update({
        where: { id: cargo.id },
        data: { aptitudes_ingreso: aptI },
      })
      updatedIngreso++
    }

    const aptP = aptMapPeriodico.get(key)
    if (aptP) {
      await prisma.cargo.update({
        where: { id: cargo.id },
        data: { aptitudes_periodico: aptP },
      })
      updatedPeriodico++
    }
  }

  console.log(`Updated ${updatedIngreso} cargos with ingreso aptitudes`)
  console.log(`Updated ${updatedPeriodico} cargos with periodico aptitudes`)
  console.log("Done!")
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(() => prisma.$disconnect())
