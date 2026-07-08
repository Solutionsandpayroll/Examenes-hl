import { NextResponse } from "next/server"
import { generateExcel, type FormData } from "@/lib/excel-generator"

export async function POST(request: Request) {
  try {
    const data: FormData = await request.json()
    const buffer = await generateExcel(data)

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="Formato_Examen.xlsx"`,
      },
    })
  } catch (error) {
    console.error("Error generating Excel:", error)
    const msg = error instanceof Error ? error.message : "Error desconocido"
    return NextResponse.json(
      { error: `Error al generar el archivo Excel: ${msg}` },
      { status: 500 }
    )
  }
}
