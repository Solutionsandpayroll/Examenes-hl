"use client"

import { useState, useEffect, useCallback } from "react"

interface Proyecto {
  id: number
  nombre: string
  nit_empresa: string | null
  nombre_empresa: string | null
  ciudad: string | null
  centro_costos: string | null
}

interface Cargo {
  id: number
  nombre: string
  proyecto_id: number
}

interface Examen {
  id: number
  nombre: string
}

const TIPOS_EXAMEN = ["INGRESO", "PERIODICO"]

export default function Home() {
  const [proyectos, setProyectos] = useState<Proyecto[]>([])
  const [selectedProyecto, setSelectedProyecto] = useState<number>(0)
  const [cargos, setCargos] = useState<Cargo[]>([])
  const [selectedCargo, setSelectedCargo] = useState<number>(0)
  const [tipoExamen, setTipoExamen] = useState("INGRESO")
  const [examenes, setExamenes] = useState<Examen[]>([])
  const [aptitudes, setAptitudes] = useState<string>("")
  const [proyectoInfo, setProyectoInfo] = useState<Proyecto | null>(null)

  const [nombreAspirante, setNombreAspirante] = useState("")
  const [cedula, setCedula] = useState("")
  const [nombreEmpresa, setNombreEmpresa] = useState("CONSORCIO HL GISAICO")
  const [nit, setNit] = useState("900966356-0")
  const [empresaMision, setEmpresaMision] = useState("")
  const [centroCostos, setCentroCostos] = useState("")
  const [ciudadProyecto, setCiudadProyecto] = useState("")
  const [nombreIps, setNombreIps] = useState("")
  const [ciudadIps, setCiudadIps] = useState("BOGOTA")
  const [telefonoIps, setTelefonoIps] = useState("")
  const [direccionIps, setDireccionIps] = useState("")
  const [fechaExamen, setFechaExamen] = useState("")
  const [horaExamen, setHoraExamen] = useState("")
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState("")

  useEffect(() => {
    fetch("/api/proyectos")
      .then((r) => r.json())
      .then(setProyectos)
      .catch(console.error)
  }, [])

  const loadCargos = useCallback(async (proyectoId: number) => {
    const res = await fetch(`/api/proyectos/${proyectoId}/cargos`)
    const data = await res.json()
    setCargos(data.cargos || [])
    setProyectoInfo(data.proyecto || null)

    if (data.proyecto) {
      if (data.proyecto.nombre_empresa) setNombreEmpresa(data.proyecto.nombre_empresa)
      if (data.proyecto.nit_empresa) setNit(data.proyecto.nit_empresa)
      if (data.proyecto.ciudad) setCiudadProyecto(data.proyecto.ciudad)
      if (data.proyecto.centro_costos) setCentroCostos(data.proyecto.centro_costos)
    }
  }, [])

  const loadExamenes = useCallback(async (cargoId: number, tipo: string) => {
    const res = await fetch(
      `/api/cargos/${cargoId}/examenes?tipo=${tipo.toLowerCase()}`
    )
    const data = await res.json()
    setExamenes(data.examenes || [])
    setAptitudes(data.aptitudes || "")
  }, [])

  const handleProyectoChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const id = parseInt(e.target.value)
    setSelectedProyecto(id)
    setSelectedCargo(0)
    setExamenes([])
    if (id) loadCargos(id)
  }

  const handleCargoChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const id = parseInt(e.target.value)
    setSelectedCargo(id)
    setExamenes([])
    if (id) loadExamenes(id, tipoExamen)
  }

  const handleTipoChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const tipo = e.target.value
    setTipoExamen(tipo)
    if (selectedCargo) loadExamenes(selectedCargo, tipo)
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError("")
    setLoading(true)

    try {
      const res = await fetch("/api/generar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          proyecto: proyectoInfo?.nombre || "",
          tipo_examen: tipoExamen,
          tipo_examen_otro: "",
          cargo: cargos.find((c) => c.id === selectedCargo)?.nombre || "",
          nombre_aspirante: nombreAspirante,
          cedula,
          nombre_empresa: nombreEmpresa,
          nit,
          empresa_mision: empresaMision,
          centro_costos: centroCostos,
          ciudad_proyecto: ciudadProyecto,
          nombre_ips: nombreIps,
          ciudad_ips: ciudadIps,
          telefono_ips: telefonoIps,
          direccion_ips: direccionIps,
          fecha_examen: fechaExamen,
          hora_examen: horaExamen,
          examenes: examenes.map((e) => e.nombre),
          aptitudes,
        }),
      })

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}))
        throw new Error((errData as { error?: string }).error || "Error al generar el archivo")
      }

      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.href = url
      a.download = "Formato_Examen.xlsx"
      a.click()
      URL.revokeObjectURL(url)
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error desconocido")
    } finally {
      setLoading(false)
    }
  }

  return (
    <form onSubmit={handleSubmit}>
      <div className="container">
        <div className="page-title">
          <h2>Formato de Solicitud</h2>
          <p>Complete los datos para generar el archivo de exámenes médicos</p>
        </div>

        <div className="card" style={{ marginBottom: "1.5rem" }}>
          <div className="card-header">
            <h3>
              <span className="icon-dot" />
              Proyecto y Tipo de Examen
            </h3>
          </div>
          <div className="card-body">
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Proyecto</label>
                <select
                  className="form-select"
                  value={selectedProyecto}
                  onChange={handleProyectoChange}
                  required
                >
                  <option value="">Seleccione un proyecto</option>
                  {proyectos.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.nombre}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label className="form-label">Tipo de Examen</label>
                <select
                  className="form-select"
                  value={tipoExamen}
                  onChange={handleTipoChange}
                  required
                >
                  {TIPOS_EXAMEN.map((t) => (
                    <option key={t} value={t}>
                      {t === "INGRESO" ? "Ingreso" : "Periódico"}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        </div>

        <div className="card" style={{ marginBottom: "1.5rem" }}>
          <div className="card-header">
            <h3>
              <span className="icon-dot" />
              Cargo y Exámenes
            </h3>
          </div>
          <div className="card-body">
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Cargo</label>
                <select
                  className="form-select"
                  value={selectedCargo}
                  onChange={handleCargoChange}
                  required
                  disabled={!selectedProyecto}
                >
                  <option value="">Seleccione un cargo</option>
                  {cargos.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.nombre}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label className="form-label">Exámenes Médicos</label>
                <div className="exam-list">
                  {examenes.length === 0 ? (
                    <span className="empty-msg">
                      {selectedCargo
                        ? "No hay exámenes configurados para este cargo"
                        : "Seleccione proyecto y cargo"}
                    </span>
                  ) : (
                    <ul>
                      {examenes.map((ex) => (
                        <li key={ex.id}>{ex.nombre}</li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
              {aptitudes && (
                <div className="form-group" style={{ gridColumn: "1 / -1" }}>
                  <label className="form-label">Aptitudes Médicas Especiales</label>
                  <div className="exam-list">
                    <ul>
                      {aptitudes.split(",").map((a, i) => (
                        <li key={i}>{a.trim()}</li>
                      ))}
                    </ul>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="card" style={{ marginBottom: "1.5rem" }}>
          <div className="card-header">
            <h3>
              <span className="icon-dot" />
              Datos del Aspirante
            </h3>
          </div>
          <div className="card-body">
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Nombre Aspirante</label>
                <input
                  className="form-input"
                  value={nombreAspirante}
                  onChange={(e) => setNombreAspirante(e.target.value)}
                  placeholder="Carlos Andrés Quiroga"
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Cédula</label>
                <input
                  className="form-input"
                  value={cedula}
                  onChange={(e) => setCedula(e.target.value)}
                  placeholder="80.853.101"
                  required
                />
              </div>
            </div>
          </div>
        </div>

        <div className="card" style={{ marginBottom: "1.5rem" }}>
          <div className="card-header">
            <h3>
              <span className="icon-dot" />
              Datos de la Empresa
            </h3>
          </div>
          <div className="card-body">
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Nombre Empresa</label>
                <input
                  className="form-input"
                  value={nombreEmpresa}
                  onChange={(e) => setNombreEmpresa(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">NIT</label>
                <input
                  className="form-input"
                  value={nit}
                  onChange={(e) => setNit(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Empresa en Misión</label>
                <input
                  className="form-input"
                  value={empresaMision}
                  onChange={(e) => setEmpresaMision(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Centro de Costos</label>
                <input
                  className="form-input"
                  value={centroCostos}
                  onChange={(e) => setCentroCostos(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Ciudad Proyecto</label>
                <input
                  className="form-input"
                  value={ciudadProyecto}
                  onChange={(e) => setCiudadProyecto(e.target.value)}
                />
              </div>
            </div>
          </div>
        </div>

        <div className="card" style={{ marginBottom: "1.5rem" }}>
          <div className="card-header">
            <h3>
              <span className="icon-dot" />
              Datos de la IPS de Atención
            </h3>
          </div>
          <div className="card-body">
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Nombre IPS</label>
                <input
                  className="form-input"
                  value={nombreIps}
                  onChange={(e) => setNombreIps(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Ciudad</label>
                <input
                  className="form-input"
                  value={ciudadIps}
                  onChange={(e) => setCiudadIps(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Teléfono</label>
                <input
                  className="form-input"
                  value={telefonoIps}
                  onChange={(e) => setTelefonoIps(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Dirección</label>
                <input
                  className="form-input"
                  value={direccionIps}
                  onChange={(e) => setDireccionIps(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Fecha del Examen</label>
                <input
                  className="form-input"
                  value={fechaExamen}
                  onChange={(e) => setFechaExamen(e.target.value)}
                  placeholder="23 de junio de 2026"
                />
              </div>
              <div className="form-group">
                <label className="form-label">Hora</label>
                <input
                  className="form-input"
                  value={horaExamen}
                  onChange={(e) => setHoraExamen(e.target.value)}
                  placeholder="7:00 a.m."
                />
              </div>
            </div>
          </div>
        </div>

        {error && <div className="error-banner">{error}</div>}

        <div style={{ textAlign: "center", marginTop: "2rem", marginBottom: "1rem" }}>
          <button
            type="submit"
            className="btn-primary"
            disabled={loading}
            style={{ maxWidth: "400px" }}
          >
            {loading ? (
              <>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ animation: "spin 1s linear infinite" }}>
                  <circle cx="12" cy="12" r="10" strokeOpacity="0.25" />
                  <path d="M12 2a10 10 0 0 1 10 10" strokeLinecap="round" />
                </svg>
                Generando...
              </>
            ) : (
              <>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                Generar Formato Excel
              </>
            )}
          </button>
        </div>
      </div>
    </form>
  )
}
