import "dotenv/config"
import { PrismaClient } from "../src/generated/prisma/client"
import { PrismaPg } from "@prisma/adapter-pg"
import ExcelJS from "exceljs"
import path from "path"

const prisma = new PrismaClient({
  adapter: new PrismaPg({ connectionString: process.env.DATABASE_URL! }),
})

const normalizeText = (text: string): string =>
  text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim()

const normalizeExamName = (name: string): string => {
  const normalized = normalizeText(name)

  if (normalized.includes("DROGAS EN ORINA") || normalized.includes("PRUEBA DE DROGAS"))
    return "PRUEBA DE DROGAS EN ORINA (PANEL X 5 DROGAS)"
  if (normalized.includes("EXAMEN MEDICO") || normalized.includes("OSTEOMUSCULAR"))
    return "EXAMEN MEDICO OCUPACIONAL OSTEOMUSCULAR"
  if (normalized.includes("CUADRO HEMATICO"))
    return "CUADRO HEMATICO"
  if (normalized.includes("GLICEMIA EN AYUNAS") || normalized === "GLICEMIA")
    return "GLICEMIA EN AYUNAS"
  if (normalized.includes("PERFIL LIPIDICO"))
    return "PERFIL LIPIDICO COMPLETO"
  if (normalized.includes("TRANSAMINASAS") || normalized.includes("ALAT"))
    return "TRANSAMINASAS (ALAT/ASAT)"
  if (normalized === "MICROALBUMINURIA") return "MICROALBUMINURIA"
  if (normalized.includes("AUDIOMETR")) return "AUDIOMETRIA"
  if (normalized.includes("OPTOMETR")) return "OPTOMETRIA"
  if (normalized.includes("PARCIAL ORINA") || normalized.includes("PARCIAL DE ORINA"))
    return "PARCIAL DE ORINA"
  if (normalized.includes("CAMPIMETRIA")) return "CAMPIMETRIA POR CONFRONTACION"
  if (normalized.includes("ESPIROMETR")) return "ESPIROMETRIA"
  if (normalized.includes("RX DE TORAX") || normalized.includes("RX TORAX"))
    return "RX DE TORAX AP Y LATERAL"
  if (normalized.includes("ELECTROCARDIOGRAMA") || normalized.includes("EKG"))
    return "ELECTROCARDIOGRAMA (MAYORES DE 45)"
  if (normalized.includes("NEUROPSICOLOGICA") || normalized.includes("PRUEBAS NEUROPSICOLOGICA"))
    return "PRUEBA NEUROPSICOLOGICA PARA ALTURAS"
  if (normalized.includes("PSICOSENSOMETRICA")) return "PRUEBAS PSICOSENSOMETRICAS"
  if (normalized.includes("COPROLOGICO") || normalized.includes("COPROPARASIT")) return "COPROLOGICO"
  if (normalized.includes("PRUEBA DE EMBARAZO") || normalized.includes("EMBARAZO"))
    return "PRUEBA DE EMBARAZO"
  if (normalized.includes("VACUNA") && (normalized.includes("TETANO") || normalized.includes("TETAN")))
    return "VACUNA TETANO"
  if (normalized.includes("VACUN") && normalized.includes("FIEBRE AMARILLA"))
    return "VACUNA FIEBRE AMARILLA"
  if (normalized.includes("ELECTROCARDIOGRAMA"))
    return "ELECTROCARDIOGRAMA (MAYORES DE 45)"

  return normalized
}

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
  console.log("Cleaning database...")
  await prisma.cargoExamenPeriodico.deleteMany()
  await prisma.cargoExamen.deleteMany()
  await prisma.cargo.deleteMany()
  await prisma.examen.deleteMany()
  await prisma.proyecto.deleteMany()

  const dataDir = path.join(process.cwd(), "prisma", "data")
  const entradaPath = path.join(dataDir, "examenes-medicos-entrada.xlsx")
  const periodicosPath = path.join(dataDir, "examenes-medicos-periodicos.xlsx")

  console.log("Reading Excel files...")
  const entradaRows = await readExcelRows(entradaPath)
  const periodicosRows = await readExcelRows(periodicosPath)

  const allCargoExams = new Set<string>()
  const allCargoExamsPeriodico = new Set<string>()

  for (const row of entradaRows) {
    const proyectoSiga = row["Proyecto"] || row["PROYECTO"]
    const cargoSigla = row["Título"] || row["TITULO"] || row["Title"] || row["TITLE"]
    const examenesStr = row["Examenes"] || row["EXAMENES"]
    if (!proyectoSiga || !cargoSigla) continue

    const proyectoNombre = proyectoSiga.trim()
    const cargoNombre = cargoSigla.trim()
    const examenes = examenesStr
      ? examenesStr.split(",").map((e: string) => normalizeExamName(e.trim())).filter(Boolean)
      : []

    for (const exam of examenes) {
      allCargoExams.add(`${proyectoNombre}|${cargoNombre}|${exam}`)
    }
  }

  for (const row of periodicosRows) {
    const proyectoPr = row["Proyecto_pr"] || row["PROYECTO_PR"]
    const proyectoSiga = row["Proyecto"] || row["PROYECTO"] || proyectoPr
    const cargoSigla = row["Title"] || row["TITLE"] || row["Título"] || row["TITULO"]
    const examenesStr = row["Examenes_Periodicos"] || row["EXAMENES_PERIODICOS"]
    if (!proyectoSiga || !cargoSigla) continue

    const proyectoNombre = proyectoSiga.trim()
    const cargoNombre = cargoSigla.trim()
    const examenes = examenesStr
      ? examenesStr.split(",").map((e: string) => normalizeExamName(e.trim())).filter(Boolean)
      : []

    for (const exam of examenes) {
      allCargoExamsPeriodico.add(`${proyectoNombre}|${cargoNombre}|${exam}`)
    }
  }

  // Collect unique values
  const uniqueProyectos = new Set<string>()
  const uniqueCargos = new Set<string>() // formato: "proyecto|cargo"
  const uniqueExamenes = new Set<string>()

  for (const ce of allCargoExams) {
    const [proy, cargo, exam] = ce.split("|")
    uniqueProyectos.add(proy)
    uniqueCargos.add(`${proy}|${cargo}`)
    uniqueExamenes.add(exam)
  }
  for (const ce of allCargoExamsPeriodico) {
    const [proy, cargo, exam] = ce.split("|")
    uniqueProyectos.add(proy)
    uniqueCargos.add(`${proy}|${cargo}`)
    uniqueExamenes.add(exam)
  }

  // BULK INSERT proyectos
  console.log(`Inserting ${uniqueProyectos.size} proyectos...`)
  await prisma.proyecto.createMany({
    data: Array.from(uniqueProyectos).map((n) => ({ nombre: n })),
    skipDuplicates: true,
  })

  // Load projects to get IDs
  const proyectosDb = await prisma.proyecto.findMany()
  const proyectoMap = new Map(proyectosDb.map((p) => [p.nombre, p.id]))

  // BULK INSERT cargos
  console.log(`Inserting ${uniqueCargos.size} cargos...`)
  const cargoRecords: { nombre: string; proyecto_id: number }[] = []
  for (const pc of uniqueCargos) {
    const [proy, cargo] = pc.split("|")
    const proyectoId = proyectoMap.get(proy)
    if (proyectoId) {
      cargoRecords.push({ nombre: cargo, proyecto_id: proyectoId })
    }
  }

  // chunked createMany to avoid huge arrays
  const chunkSize = 500
  for (let i = 0; i < cargoRecords.length; i += chunkSize) {
    const chunk = cargoRecords.slice(i, i + chunkSize)
    await prisma.cargo.createMany({
      data: chunk,
      skipDuplicates: true,
    })
    console.log(`  Inserted ${Math.min(i + chunkSize, cargoRecords.length)}/${cargoRecords.length} cargos`)
  }

  // Load cargos to get IDs
  const cargosDb = await prisma.cargo.findMany({
    include: { proyecto: true },
  })
  const cargoMap = new Map<string, number>()
  for (const c of cargosDb) {
    cargoMap.set(`${c.proyecto.nombre}|${c.nombre}`, c.id)
  }

  // BULK INSERT examenes
  console.log(`Inserting ${uniqueExamenes.size} examenes...`)
  await prisma.examen.createMany({
    data: Array.from(uniqueExamenes).map((n) => ({ nombre: n })),
    skipDuplicates: true,
  })

  // Load examenes to get IDs
  const examenesDb = await prisma.examen.findMany()
  const examenMap = new Map(examenesDb.map((e) => [e.nombre, e.id]))

  // BULK INSERT cargo-examen relations
  console.log("Inserting cargo-examen ingreso relations...")
  const ingresoRecords: { cargo_id: number; examen_id: number }[] = []
  for (const ce of allCargoExams) {
    const [proy, cargo, exam] = ce.split("|")
    const cargoId = cargoMap.get(`${proy}|${cargo}`)
    const examenId = examenMap.get(exam)
    if (cargoId && examenId) {
      ingresoRecords.push({ cargo_id: cargoId, examen_id: examenId })
    }
  }
  for (let i = 0; i < ingresoRecords.length; i += chunkSize) {
    const chunk = ingresoRecords.slice(i, i + chunkSize)
    await prisma.cargoExamen.createMany({
      data: chunk,
      skipDuplicates: true,
    })
  }
  console.log(`  Inserted ${ingresoRecords.length} ingreso relations`)

  console.log("Inserting cargo-examen periodico relations...")
  const periodicoRecords: { cargo_id: number; examen_id: number }[] = []
  for (const ce of allCargoExamsPeriodico) {
    const [proy, cargo, exam] = ce.split("|")
    const cargoId = cargoMap.get(`${proy}|${cargo}`)
    const examenId = examenMap.get(exam)
    if (cargoId && examenId) {
      periodicoRecords.push({ cargo_id: cargoId, examen_id: examenId })
    }
  }
  for (let i = 0; i < periodicoRecords.length; i += chunkSize) {
    const chunk = periodicoRecords.slice(i, i + chunkSize)
    await prisma.cargoExamenPeriodico.createMany({
      data: chunk,
      skipDuplicates: true,
    })
  }
  console.log(`  Inserted ${periodicoRecords.length} periodico relations`)

  console.log("Seed complete!")
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(() => prisma.$disconnect())
