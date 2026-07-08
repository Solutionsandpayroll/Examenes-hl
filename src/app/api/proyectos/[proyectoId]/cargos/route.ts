import { NextResponse } from "next/server"
import { prisma } from "@/lib/prisma"

export async function GET(
  _request: Request,
  { params }: { params: Promise<{ proyectoId: string }> }
) {
  const { proyectoId } = await params
  const id = parseInt(proyectoId)

  const cargos = await prisma.cargo.findMany({
    where: { proyecto_id: id },
    orderBy: { nombre: "asc" },
  })

  const proyecto = await prisma.proyecto.findUnique({ where: { id } })

  return NextResponse.json({ cargos, proyecto })
}
