import React from 'react';
import {
  Printer,
  Camera,
  Clock,
  CheckCircle2,
} from 'lucide-react';

interface HeaderProps {
  lastUpdatedText: string;
  onPrint: () => void;
  onOpenCapture: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  lastUpdatedText,
  onPrint,
  onOpenCapture,
}) => {
  return (
    <header
      id="main-header"
      className="bg-[#691C32] text-white border-b-4 border-[#D4C19C] shadow-sm print:bg-white print:text-black print:border-b-2 print:border-gray-800"
    >
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-5">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          {/* Logo & Titles */}
          <div className="flex items-start sm:items-center gap-3.5">
            <div className="h-16 w-32 flex items-center justify-center shrink-0">
              <img
                src="https://imssbienestar.gob.mx/assets/img/imb_b.svg"
                alt="IMSS-Bienestar"
                className="max-h-full max-w-full object-contain"
              />
            </div>

            <div>
              <div className="flex items-center gap-2">
                <span className="hidden sm:inline-flex items-center gap-1 text-[11px] font-medium text-emerald-300 print:text-emerald-800 bg-black/20 px-2 py-0.5 rounded">
                  <CheckCircle2 className="w-3 h-3" />
                  Datos Validados
                </span>
              </div>
              <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-white print:text-black font-['Montserrat',sans-serif]">
                Reporte de Inventario
              </h1>
              <p className="text-xs sm:text-sm text-white/80 print:text-gray-700 mt-0.5">
                Inventario de unidades médicas por entidad federativa
              </p>
            </div>
          </div>

          {/* Date and Action Controls */}
          <div className="flex flex-col sm:flex-row sm:items-center gap-3 lg:gap-4 print:hidden">
            {/* Dynamic Timestamp */}
            <div className="flex items-center gap-1.5 text-xs text-white/85 bg-black/20 px-3 py-1.5 rounded border border-white/10">
              <Clock className="w-3.5 h-3.5 text-[#D4C19C]" />
              <span>
                Última actualización:{' '}
                <strong className="font-semibold text-white">{lastUpdatedText}</strong>
              </span>
            </div>

            {/* Action Buttons */}
            <div className="flex items-center gap-2">
              {/* Print / PDF Button */}
              <button
                id="btn-print"
                type="button"
                onClick={onPrint}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded bg-white/10 hover:bg-white/20 text-white border border-white/20 transition-colors focus:outline-none focus:ring-2 focus:ring-[#D4C19C]"
                title="Imprimir o exportar como documento PDF"
              >
                <Printer className="w-3.5 h-3.5 text-[#D4C19C]" />
                <span>Imprimir / PDF</span>
              </button>

              {/* Capture Report Dropdown */}
              <div className="relative">
                <button
                  id="btn-capture-menu"
                  type="button"
                  onClick={onOpenCapture}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded bg-[#D4C19C] hover:bg-[#c4b08a] text-[#10312B] font-semibold transition-colors focus:outline-none focus:ring-2 focus:ring-white"
                  title="Capturar reporte como imagen PNG"
                >
                  <Camera className="w-3.5 h-3.5" />
                  <span>Capturar reporte</span>
                </button>
              </div>

            </div>
          </div>
        </div>

        {/* Print-only institutional header subtitle */}
        <div className="hidden print:flex items-center justify-between text-xs text-gray-600 border-t border-gray-300 pt-2 mt-2">
          <span>Coordinación Nacional de Abasto y Equipamiento Médico</span>
          <span>Fecha de emisión: {lastUpdatedText}</span>
        </div>
      </div>
    </header>
  );
};
