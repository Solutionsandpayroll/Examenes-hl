import { NextResponse } from "next/server"
import { prisma } from "@/lib/prisma"

export async function GET(
  request: Request,
  { params }: { params: Promise<{ cargoId: string }> }
) {
  const { cargoId } = await params
  const id = parseInt(cargoId)

  const { searchParams } = new URL(request.url)
  const tipo = searchParams.get("tipo") // "ingreso" o "periodico"

  const cargo = await prisma.cargo.findUnique({ where: { id } })
  if (!cargo) return NextResponse.json({ error: "Cargo not found" }, { status: 404 })

  let examenes
  if (tipo === "periodico") {
    const relations = await prisma.cargoExamenPeriodico.findMany({
      where: { cargo_id: id },
      include: { examen: true },
      orderBy: { examen: { nombre: "asc" } },
    })
    examenes = relations.map((r) => r.examen)
  } else {
    const relations = await prisma.cargoExamen.findMany({
      where: { cargo_id: id },
      include: { examen: true },
      orderBy: { examen: { nombre: "asc" } },
    })
    examenes = relations.map((r) => r.examen)
  }

  const aptitudes =
    tipo === "periodico" ? cargo.aptitudes_periodico : cargo.aptitudes_ingreso

  return NextResponse.json({ examenes, aptitudes })
}
