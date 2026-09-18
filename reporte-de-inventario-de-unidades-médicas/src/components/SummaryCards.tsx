import React from 'react';
import { NationalSummary } from '../types';
import { formatNumber, formatPercent } from '../utils/formatters';
import { Target, ClipboardCheck, TrendingUp, ShieldCheck } from 'lucide-react';
import { getSemaforoStyles } from '../config/thresholds';

interface SummaryCardsProps {
  summary: NationalSummary;
  filteredEntityName?: string;
}

export const SummaryCards: React.FC<SummaryCardsProps> = ({
  summary,
  filteredEntityName,
}) => {
  const avanceStyles = getSemaforoStyles(summary.porcentajeAvance);
  const completoStyles = getSemaforoStyles(summary.porcentajeCompleto);

  return (
    <section
      id="capture-summary"
      aria-labelledby="heading-vista-general"
      className="bg-white rounded-lg border border-gray-200 shadow-xs p-4 sm:p-5 print:border-gray-300 print:shadow-none"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-3 mb-4 border-b border-gray-100 gap-2">
        <div>
          <h2
            id="heading-vista-general"
            className="text-base sm:text-lg font-bold text-[#10312B] font-['Montserrat',sans-serif] flex items-center gap-2"
          >
            <span>Vista General</span>
            {filteredEntityName && (
              <span className="text-xs font-medium text-[#691C32] bg-[#691C32]/10 px-2 py-0.5 rounded">
                Filtrado por: {filteredEntityName}
              </span>
            )}
          </h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Métricas ejecutivas consolidadas del avance en el reporte de inventarios
          </p>
        </div>

        <div className="hidden sm:flex items-center gap-3 text-xs text-gray-500">
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-[#137537]" />
            <span>≥70% Óptimo</span>
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-[#E5B824]" />
            <span>50-69% Regular</span>
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-[#C82333]" />
            <span>&lt;50% Prioritario</span>
          </span>
        </div>
      </div>

      {/* 4 Executive KPI Cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
        {/* Card 1: Meta de CLUES */}
        <div
          id="kpi-meta"
          className="bg-gray-50/70 rounded-md border border-gray-200/80 p-3.5 sm:p-4 flex flex-col justify-between hover:border-gray-300 transition-colors"
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold uppercase tracking-wider text-gray-500">
              Meta de CLUES
            </span>
            <span className="p-1.5 rounded bg-gray-200/60 text-gray-700">
              <Target className="w-4 h-4" />
            </span>
          </div>
          <div className="mt-2">
            <div className="text-2xl sm:text-3xl font-extrabold text-[#10312B] tracking-tight tabular-nums">
              {formatNumber(summary.metaTotal)}
            </div>
            <p className="text-[11px] text-gray-500 mt-0.5">
              Unidades médicas programadas
            </p>
          </div>
        </div>

        {/* Card 2: CLUES con inventario */}
        <div
          id="kpi-con-inventario"
          className="bg-gray-50/70 rounded-md border border-gray-200/80 p-3.5 sm:p-4 flex flex-col justify-between hover:border-gray-300 transition-colors"
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold uppercase tracking-wider text-gray-500">
              CLUES con inventario
            </span>
            <span className="p-1.5 rounded bg-[#691C32]/10 text-[#691C32]">
              <ClipboardCheck className="w-4 h-4" />
            </span>
          </div>
          <div className="mt-2">
            <div className="text-2xl sm:text-3xl font-extrabold text-[#691C32] tracking-tight tabular-nums">
              {formatNumber(summary.conInventarioTotal)}
            </div>
            <p className="text-[11px] text-gray-500 mt-0.5">
              Unidades con registro activo
            </p>
          </div>
        </div>

        {/* Card 3: % de avance */}
        <div
          id="kpi-avance"
          className={`rounded-md border p-3.5 sm:p-4 flex flex-col justify-between transition-colors ${avanceStyles.bgColor} ${avanceStyles.borderColor}`}
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold uppercase tracking-wider text-gray-600">
              % de avance
            </span>
            <span className="p-1.5 rounded bg-white/70 text-gray-700">
              <TrendingUp className="w-4 h-4" />
            </span>
          </div>
          <div className="mt-2">
            <div className={`text-2xl sm:text-3xl font-extrabold tracking-tight tabular-nums ${avanceStyles.textColor}`}>
              {formatPercent(summary.porcentajeAvance)}
            </div>
            <div className="flex items-center gap-1.5 mt-0.5 text-[11px] font-medium text-gray-600">
              <span className={`w-2 h-2 rounded-full ${avanceStyles.dotColor}`} />
              <span>Avance global de captura</span>
            </div>
          </div>
        </div>

        {/* Card 4: % completo */}
        <div
          id="kpi-completo"
          className={`rounded-md border p-3.5 sm:p-4 flex flex-col justify-between transition-colors ${completoStyles.bgColor} ${completoStyles.borderColor}`}
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold uppercase tracking-wider text-gray-600">
              % completo
            </span>
            <span className="p-1.5 rounded bg-white/70 text-[#10312B]">
              <ShieldCheck className="w-4 h-4" />
            </span>
          </div>
          <div className="mt-2">
            <div className={`text-2xl sm:text-3xl font-extrabold tracking-tight tabular-nums ${completoStyles.textColor}`}>
              {formatPercent(summary.porcentajeCompleto)}
            </div>
            <div className="flex items-center gap-1.5 mt-0.5 text-[11px] font-medium text-gray-600">
              <span className={`w-2 h-2 rounded-full ${completoStyles.dotColor}`} />
              <span>{formatNumber(summary.inventarioCompletoTotal)} unidades 100% completas</span>
            </div>
          </div>
        </div>
      </div>

      {/* Sub-bar showing Catalogs Breakdown */}
      <div className="mt-3 pt-3 border-t border-gray-100 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 text-xs text-gray-600">
        <div className="flex items-center justify-between bg-gray-50 px-3 py-1.5 rounded border border-gray-100">
          <span className="text-gray-500">CLUES Medicamentos (010/040):</span>
          <span className="font-semibold text-gray-800 tabular-nums">
            {formatNumber(summary.medicamentosTotal)}{' '}
            <span className="font-normal text-gray-400">
              ({formatPercent((summary.medicamentosTotal / (summary.metaTotal || 1)) * 100)})
            </span>
          </span>
        </div>
        <div className="flex items-center justify-between bg-gray-50 px-3 py-1.5 rounded border border-gray-100">
          <span className="text-gray-500">CLUES Mat. de Curación (060):</span>
          <span className="font-semibold text-gray-800 tabular-nums">
            {formatNumber(summary.materialCuracionTotal)}{' '}
            <span className="font-normal text-gray-400">
              ({formatPercent((summary.materialCuracionTotal / (summary.metaTotal || 1)) * 100)})
            </span>
          </span>
        </div>
        <div className="flex items-center justify-between bg-gray-50 px-3 py-1.5 rounded border border-gray-100 sm:col-span-2 lg:col-span-1">
          <span className="text-gray-500">Inventario Completo (Ambos):</span>
          <span className="font-semibold text-[#10312B] tabular-nums">
            {formatNumber(summary.inventarioCompletoTotal)}{' '}
            <span className="font-normal text-gray-400">
              ({formatPercent((summary.inventarioCompletoTotal / (summary.metaTotal || 1)) * 100)})
            </span>
          </span>
        </div>
      </div>
    </section>
  );
};
