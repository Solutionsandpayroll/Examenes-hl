import type { Metadata } from "next"
import "./globals.css"

export const metadata: Metadata = {
  title: "Exámenes Médicos - Solutions & Payroll",
  description: "Solicitud de Servicio de Exámenes Médicos Ocupacionales",
  icons: { icon: "/logo-syp.png" },
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="es">
      <body>
        <div className="app-shell">
          <header className="app-header">
            <div className="container">
              <div className="header-content">
                <div className="logo-container">
                  <div className="logo">
                    <img
                      src="/logo-syp.png"
                      alt="Solutions & Payroll Logo"
                      width={60}
                      height={60}
                    />
                  </div>
                  <div className="header-text">
                    <h1>Solutions &amp; Payroll</h1>
                    <p className="subtitle">Solicitud - Exámenes Médicos Ocupacionales</p>
                  </div>
                </div>
                <div className="welcome-box">
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                    <circle cx="12" cy="7" r="4" />
                  </svg>
                  <span>Bienvenido Usuario</span>
                </div>
              </div>
            </div>
          </header>

          <main className="main-content">{children}</main>

          <footer className="app-footer">
            <div className="container">
              <p>&copy; {new Date().getFullYear()} Solutions &amp; Payroll — Exámenes Médicos Ocupacionales</p>
            </div>
          </footer>
        </div>
      </body>
    </html>
  )
}
