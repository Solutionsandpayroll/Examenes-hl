import { NextResponse } from "next/server"
import { prisma } from "@/lib/prisma"

export async function GET() {
  const proyectos = await prisma.proyecto.findMany({
    orderBy: { nombre: "asc" },
  })
  return NextResponse.json(proyectos)
}
