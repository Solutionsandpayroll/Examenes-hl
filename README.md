# Exámenes Médicos Ocupacionales

Aplicación web para generar el formato de solicitud de exámenes médicos ocupacionales de **Solutions & Payroll**. Permite diligenciar los datos del aspirante, empresa e IPS, y genera un archivo Excel con el formato oficial, incluyendo los exámenes médicos y aptitudes correspondientes según el cargo y tipo de examen seleccionado.

## Tecnologías

| Capa | Tecnología |
|------|-----------|
| Frontend | Next.js 16 + React 19 + TypeScript |
| Estilos | CSS vanilla con custom properties (paleta corporativa) |
| Backend | Next.js API Routes |
| Base de datos | PostgreSQL en [Neon](https://neon.tech) |
| ORM | Prisma 7 |
| Excel | Manipulación directa del XML interno del `.xlsx` (JSZip + DOM) |

## Estructura del proyecto

```
web/
├── prisma/
│   ├── schema.prisma          # Modelos: Proyecto, Cargo, Examen, CargoExamen, CargoExamenPeriodico
│   ├── seed.ts                # Importa datos de los Excel de mapeo a la DB
│   └── seed-aptitudes.ts      # Importa datos de aptitudes a la DB
├── public/
│   ├── formato-examen.xlsx    # Plantilla base del formato Excel
│   └── logo-syp.png           # Logo corporativo
├── src/
│   ├── app/
│   │   ├── api/
│   │   │   ├── proyectos/route.ts                  # GET → lista de proyectos
│   │   │   ├── proyectos/[proyectoId]/cargos/route.ts  # GET → cargos + info del proyecto
│   │   │   ├── cargos/[cargoId]/examenes/route.ts      # GET → exámenes y aptitudes (?tipo=ingreso|periodico)
│   │   │   └── generar/route.ts                    # POST → genera y descarga el Excel
│   │   ├── globals.css        # Variables CSS, animaciones, estilos de layout/cards/forms
│   │   ├── layout.tsx         # Layout principal (header, main, footer)
│   │   └── page.tsx           # Formulario principal (client component)
│   ├── lib/
│   │   ├── prisma.ts          # Singleton del cliente Prisma
│   │   └── excel-generator.ts # Motor de generación del Excel
│   └── generated/prisma/      # Cliente Prisma generado (no se commitea)
├── .env                       # DATABASE_URL (no se commitea)
├── package.json
└── tsconfig.json
```

## Cómo funciona

### 1. Base de datos (una sola vez)

Los datos de mapeo provienen de 3 archivos Excel:

- `Examenes medicos entrada.xlsx` → Proyecto, Cargo, Exámenes de ingreso, Aptitudes
- `Examenes medicos periodicos.xlsx` → Proyecto, Cargo, Exámenes periódicos, Aptitudes
- `Formato Examen.xlsx` → Plantilla del formato final

Los scripts `prisma/seed.ts` y `prisma/seed-aptitudes.ts` leen estos Excel, normalizan los nombres de exámenes y aptitudes, y los insertan en PostgreSQL (Neon). Esto evita tener que subir los Excel cada vez.

### 2. Flujo del formulario

1. Usuario selecciona **Proyecto** → se cargan los cargos y se precargan NIT, empresa, ciudad y centro de costos
2. Usuario selecciona **Tipo de Examen** (Ingreso o Periódico)
3. Usuario selecciona **Cargo** → se muestran automáticamente los exámenes y aptitudes que corresponden
4. Usuario completa los demás datos (aspirante, IPS, fecha, hora)
5. Clic en **"Generar Formato Excel"**

### 3. Generación del Excel

El motor `excel-generator.ts`:

1. Carga el template `Formato Examen.xlsx` como un ZIP
2. Usa un parser DOM (`@xmldom/xmldom`) para modificar el `sheet1.xml` de forma segura (sin regex)
3. Rellena las celdas de datos con los valores del formulario, preservando los estilos originales (`s="96"`, `s="93"`, etc.)
4. Marca los checkboxes de exámenes que corresponden (`checked="Checked"` en los `ctrlProp*.xml`)
5. Pone "X" en las celdas de aptitudes que aplican (D50, D51, F50, F51)
6. Genera el archivo `.xlsx` final

### Normalización de exámenes

Los nombres en los Excel de mapeo varían (ej: "CUADRO HEMÁTICO" vs "CUADRO HEMATICO", "VACUNACIÓN TETANO" vs "VACUNA TETANO"). El seed los normaliza (quita tildes, unifica nombres) para que coincidan con los 28 checkboxes del template.

### Normalización de aptitudes

Las aptitudes en los Excel usan nombres como "TRABAJO EN ALTURAS", "TRABAJO EN ESPACIOS CONFINADOS", etc. El generador las mapea a las celdas correspondientes del template sin importar tildes o mayúsculas.

## Configuración inicial

### Requisitos

- Node.js 22+
- Una base de datos PostgreSQL en [Neon](https://neon.tech) (gratuita)

### Instalación

```bash
cd web
npm install
```

### Variables de entorno

Crear `.env` con la URL de conexión a Neon:

```
DATABASE_URL="postgresql://usuario:password@host/neondb?sslmode=require"
```

### Base de datos

```bash
# Crear las tablas
npx prisma db push

# Generar el cliente Prisma
npx prisma generate

# Poblar la base de datos con los datos de los Excel
npm run db:seed          # Proyectos, cargos, exámenes
npx tsx prisma/seed-aptitudes.ts   # Aptitudes
```

### Desarrollo

```bash
npm run dev
```

Abre [http://localhost:3000](http://localhost:3000).

### Producción

```bash
npm run build
npm start
```

## API Endpoints

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/api/proyectos` | Lista todos los proyectos |
| GET | `/api/proyectos/[id]/cargos` | Cargos e info de un proyecto |
| GET | `/api/cargos/[id]/examenes?tipo=ingreso\|periodico` | Exámenes y aptitudes de un cargo |
| POST | `/api/generar` | Genera y descarga el Excel (body JSON) |

## Paleta de colores

| Variable | Hex |
|----------|-----|
| `--primary` | `#1e3a8a` |
| `--primary-dark` | `#1e40af` |
| `--background` | `#f8fafc` |
| `--surface` | `#ffffff` |
| `--text-primary` | `#0f172a` |
| `--text-secondary` | `#64748b` |
| `--border` | `#e2e8f0` |
