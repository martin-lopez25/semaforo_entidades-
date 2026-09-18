import React, { useState, useMemo, useRef } from 'react';
import { EntityInventory } from '../types';
import { calculateAvance, calculateCompleto } from '../utils/calculations';
import { formatPercent, formatNumber } from '../utils/formatters';
import { getSemaforoStyles, SEMAFORO_THRESHOLDS } from '../config/thresholds';
import {
  ArrowDownUp,
  BarChart2,
  BarChart3,
  AlignLeft,
  Info,
  Maximize2,
  Minimize2,
} from 'lucide-react';

interface InventoryChartProps {
  data: EntityInventory[];
  selectedEntity?: string;
  onSelectEntity?: (entity: string) => void;
}

type ChartOrientation = 'vertical' | 'horizontal';
type ChartSort = 'desc' | 'asc' | 'alpha';
type ChartDisplayLimit = 'all' | 'top10' | 'bottom10';
type ChartMetric = 'avance' | 'completo' | 'compare';

export const InventoryChart: React.FC<InventoryChartProps> = ({
  data,
  selectedEntity,
  onSelectEntity,
}) => {
  // Vertical is default as explicitly requested by user ("la grafica es mas asi en vertical")
  const [orientation, setOrientation] = useState<ChartOrientation>('vertical');
  const [sortOrder, setSortOrder] = useState<ChartSort>('desc');
  const [displayLimit, setDisplayLimit] = useState<ChartDisplayLimit>('all');
  const [metric, setMetric] = useState<ChartMetric>('avance');
  const [isFitMode, setIsFitMode] = useState<boolean>(false);
  const [hoveredEntity, setHoveredEntity] = useState<string | null>(null);

  const scrollContainerRef = useRef<HTMLDivElement>(null);

  // Process data with metrics and sorting
  const processedData = useMemo(() => {
    const items = data.map((item) => {
      const avance = calculateAvance(item.cluesConInventario, item.metaClues);
      const porcCompleto = calculateCompleto(item.inventarioCompleto, item.metaClues);
      return {
        ...item,
        avance,
        porcCompleto,
      };
    });

    // Sort items
    if (sortOrder === 'desc') {
      const sortVal = (x: typeof items[0]) => (metric === 'completo' ? x.porcCompleto : x.avance);
      items.sort((a, b) => sortVal(b) - sortVal(a));
    } else if (sortOrder === 'asc') {
      const sortVal = (x: typeof items[0]) => (metric === 'completo' ? x.porcCompleto : x.avance);
      items.sort((a, b) => sortVal(a) - sortVal(b));
    } else {
      items.sort((a, b) => a.entidad.localeCompare(b.entidad, 'es'));
    }

    // Limit if requested
    if (displayLimit === 'top10') {
      const sortedDesc = [...items].sort((a, b) => b.avance - a.avance);
      return sortedDesc.slice(0, 10);
    }
    if (displayLimit === 'bottom10') {
      const sortedAsc = [...items].sort((a, b) => a.avance - b.avance);
      return sortedAsc.slice(0, 10);
    }

    return items;
  }, [data, sortOrder, displayLimit, metric]);

  // General national average for current processed items
  const averageAvance = useMemo(() => {
    if (processedData.length === 0) return 0;
    const totalCon = processedData.reduce((acc, curr) => acc + curr.cluesConInventario, 0);
    const totalMeta = processedData.reduce((acc, curr) => acc + curr.metaClues, 0);
    return totalMeta > 0 ? (totalCon / totalMeta) * 100 : 0;
  }, [processedData]);

  const activeHoveredItem = useMemo(() => {
    if (!hoveredEntity) return null;
    return processedData.find((item) => item.entidad === hoveredEntity) || null;
  }, [hoveredEntity, processedData]);

  return (
    <section
      id="capture-chart"
      aria-labelledby="heading-grafica-avance"
      className="bg-white rounded-lg border border-gray-200 shadow-xs p-4 sm:p-5 print:border-gray-300 print:shadow-none transition-all"
    >
      {/* Header of the Chart Section */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between pb-3 mb-4 border-b border-gray-100 gap-3">
        <div>
          <div className="flex items-center gap-2">
            <BarChart3 className="w-5 h-5 text-[#691C32]" />
            <h2
              id="heading-grafica-avance"
              className="text-base sm:text-lg font-bold text-[#10312B] font-['Montserrat',sans-serif]"
            >
              Gráfica de Avance de Inventario por Entidad Federativa
            </h2>
            <span className="hidden sm:inline-block px-2 py-0.5 rounded text-[10px] font-bold bg-[#691C32]/10 text-[#691C32] uppercase tracking-wide">
              {orientation === 'vertical' ? 'Columnas Verticales' : 'Barras Horizontales'}
            </span>
          </div>
          <p className="text-xs text-gray-500 mt-0.5">
            Avance de captura (%) sobre la meta de CLUES con umbrales normativos oficiales (70% y 90%)
          </p>
        </div>

        {/* Controls Toolbar */}
        <div className="flex flex-wrap items-center gap-2 print:hidden">
          {/* Orientation Toggle: Vertical (Default) vs Horizontal */}
          <div className="inline-flex rounded border border-gray-200 p-0.5 bg-gray-50 text-xs shadow-2xs">
            <button
              type="button"
              onClick={() => setOrientation('vertical')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded font-semibold transition-all ${
                orientation === 'vertical'
                  ? 'bg-[#691C32] text-white shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
              title="Ver gráfica en columnas verticales"
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span>Vertical</span>
            </button>
            <button
              type="button"
              onClick={() => setOrientation('horizontal')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded font-semibold transition-all ${
                orientation === 'horizontal'
                  ? 'bg-[#691C32] text-white shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
              title="Ver gráfica en barras horizontales"
            >
              <AlignLeft className="w-3.5 h-3.5" />
              <span>Horizontal</span>
            </button>
          </div>

          {/* Metric Selector */}
          <div className="inline-flex rounded border border-gray-200 p-0.5 bg-gray-50 text-xs">
            <button
              type="button"
              onClick={() => setMetric('avance')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                metric === 'avance'
                  ? 'bg-white text-[#10312B] font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              % Avance
            </button>
            <button
              type="button"
              onClick={() => setMetric('completo')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                metric === 'completo'
                  ? 'bg-white text-[#10312B] font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              % Completo
            </button>
            <button
              type="button"
              onClick={() => setMetric('compare')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                metric === 'compare'
                  ? 'bg-white text-[#10312B] font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Comparativo
            </button>
          </div>

          {/* Display Limit Filter */}
          <div className="inline-flex rounded border border-gray-200 p-0.5 bg-gray-50 text-xs">
            <button
              type="button"
              onClick={() => setDisplayLimit('all')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                displayLimit === 'all'
                  ? 'bg-white text-[#10312B] font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Todas ({data.length})
            </button>
            <button
              type="button"
              onClick={() => setDisplayLimit('top10')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                displayLimit === 'top10'
                  ? 'bg-white text-emerald-800 font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Top 10
            </button>
            <button
              type="button"
              onClick={() => setDisplayLimit('bottom10')}
              className={`px-2 py-1 rounded font-medium transition-colors ${
                displayLimit === 'bottom10'
                  ? 'bg-white text-rose-800 font-bold shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              10 Menor
            </button>
          </div>

          {/* Sorting */}
          <div className="flex items-center gap-1 text-xs text-gray-500">
            <ArrowDownUp className="w-3 h-3 text-gray-400" />
            <select
              value={sortOrder}
              onChange={(e) => setSortOrder(e.target.value as ChartSort)}
              className="py-1 px-2 text-xs bg-gray-50 border border-gray-200 rounded text-gray-700 cursor-pointer focus:ring-1 focus:ring-[#691C32]"
              aria-label="Orden de la gráfica"
            >
              <option value="desc">Mayor % a menor</option>
              <option value="asc">Menor % a mayor</option>
              <option value="alpha">Nombre (A-Z)</option>
            </select>
          </div>

          {/* Scroll / Fit toggle for vertical view */}
          {orientation === 'vertical' && processedData.length > 12 && (
            <button
              type="button"
              onClick={() => setIsFitMode(!isFitMode)}
              className="inline-flex items-center gap-1 px-2 py-1 rounded text-xs border border-gray-200 bg-gray-50 text-gray-600 hover:text-gray-900 hover:bg-white transition-colors"
              title={isFitMode ? 'Ver con desplazamiento horizontal cómodo' : 'Ajustar todo al ancho'}
            >
              {isFitMode ? <Maximize2 className="w-3 h-3" /> : <Minimize2 className="w-3 h-3" />}
              <span>{isFitMode ? 'Desplazar' : 'Ajustar'}</span>
            </button>
          )}
        </div>
      </div>

      {/* Floating or Top Inspector Card when hovering a column */}
      {activeHoveredItem && (
        <div className="mb-3 p-2.5 bg-[#FAF7F0] border border-[#D4C19C]/60 rounded-md flex flex-wrap items-center justify-between text-xs gap-3 animate-in fade-in duration-150">
          <div className="flex items-center gap-2">
            <span
              className="w-2.5 h-2.5 rounded-full"
              style={{ backgroundColor: getSemaforoStyles(activeHoveredItem.avance).hexColor }}
            />
            <span className="font-bold text-[#10312B] uppercase text-xs sm:text-sm">
              {activeHoveredItem.entidad}
            </span>
            <span className="text-gray-500">·</span>
            <span className="text-gray-700 font-medium">
              Meta: <strong>{formatNumber(activeHoveredItem.metaClues)} CLUES</strong>
            </span>
            <span className="text-gray-500">·</span>
            <span className="text-gray-700 font-medium">
              Con Inventario: <strong>{formatNumber(activeHoveredItem.cluesConInventario)}</strong>
            </span>
          </div>

          <div className="flex items-center gap-3">
            <div className="flex items-center gap-1.5">
              <span className="text-gray-500">% Avance:</span>
              <span
                className="font-bold px-2 py-0.5 rounded text-white text-xs"
                style={{ backgroundColor: getSemaforoStyles(activeHoveredItem.avance).hexColor }}
              >
                {formatPercent(activeHoveredItem.avance)}
              </span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="text-gray-500">% Completo:</span>
              <span className="font-bold text-gray-800">
                {formatPercent(activeHoveredItem.porcCompleto)}
              </span>
            </div>
            <div className="hidden md:flex items-center gap-2 text-[11px] text-gray-600 border-l border-[#D4C19C]/60 pl-3">
              <span>Med: <strong>{activeHoveredItem.cluesMedicamentos}</strong></span>
              <span>Mat: <strong>{activeHoveredItem.cluesMaterialCuracion}</strong></span>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 1. VERTICAL COLUMN CHART (Requested Primary Presentation) */}
      {/* ========================================================================= */}
      {orientation === 'vertical' && (
        <div className="relative">
          {/* Main Chart Canvas Frame */}
          <div className="relative border border-gray-200 rounded-lg bg-white overflow-hidden">
            {/* Chart Area with Left Y-Axis and Horizontal Gridlines */}
            <div className="flex">
              {/* Left Y-Axis Scale (0% - 100%) */}
              <div className="w-11 sm:w-13 shrink-0 flex flex-col justify-between py-6 pr-1 text-[11px] font-semibold text-gray-400 select-none border-r border-gray-200 bg-gray-50/70 text-right">
                <span className="tabular-nums">100%</span>
                <span className="tabular-nums text-emerald-700 font-bold">90%</span>
                <span className="tabular-nums">80%</span>
                <span className="tabular-nums text-amber-600 font-bold">70%</span>
                <span className="tabular-nums">60%</span>
                <span className="tabular-nums text-rose-600 font-bold">50%</span>
                <span className="tabular-nums">40%</span>
                <span className="tabular-nums">20%</span>
                <span className="tabular-nums">0%</span>
              </div>

              {/* Scrollable / Responsive Columns Area */}
              <div
                ref={scrollContainerRef}
                className={`flex-1 relative overflow-x-auto ${
                  isFitMode ? 'overflow-x-hidden' : ''
                } pb-3`}
                style={{ minHeight: '410px' }}
              >
                {/* Horizontal Reference Gridlines */}
                <div className="absolute inset-x-0 top-6 bottom-32 pointer-events-none z-0">
                  {/* 100% Line */}
                  <div className="absolute top-[0%] inset-x-0 border-b border-gray-200" />
                  
                  {/* 90% Institutional Target Line */}
                  <div className="absolute top-[10%] inset-x-0 border-b-2 border-dashed border-[#137537]/50 flex items-center justify-end pr-2">
                    <span className="text-[10px] font-bold text-[#137537] bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-300/60 shadow-2xs">
                      Meta 90%
                    </span>
                  </div>

                  {/* 80% Line */}
                  <div className="absolute top-[20%] inset-x-0 border-b border-gray-100" />

                  {/* 70% Institutional Minimum Threshold Line */}
                  <div className="absolute top-[30%] inset-x-0 border-b-2 border-dashed border-[#D97706]/60 flex items-center justify-end pr-2">
                    <span className="text-[10px] font-bold text-[#B45309] bg-amber-50 px-1.5 py-0.5 rounded border border-amber-300/60 shadow-2xs">
                      Umbral 70%
                    </span>
                  </div>

                  {/* 60% Line */}
                  <div className="absolute top-[40%] inset-x-0 border-b border-gray-100" />

                  {/* 50% Critical Line */}
                  <div className="absolute top-[50%] inset-x-0 border-b-2 border-dashed border-[#C82333]/50 flex items-center justify-end pr-2">
                    <span className="text-[10px] font-bold text-[#C82333] bg-rose-50 px-1.5 py-0.5 rounded border border-rose-300/60 shadow-2xs">
                      Crítico 50%
                    </span>
                  </div>

                  {/* 40% Line */}
                  <div className="absolute top-[60%] inset-x-0 border-b border-gray-100" />

                  {/* 20% Line */}
                  <div className="absolute top-[80%] inset-x-0 border-b border-gray-100" />

                  {/* 0% Baseline */}
                  <div className="absolute top-[100%] inset-x-0 border-b-2 border-gray-400" />
                </div>

                {/* Columns Container */}
                <div
                  className={`relative z-10 flex items-end h-[280px] mt-6 px-3 gap-2 sm:gap-3 ${
                    isFitMode
                      ? 'w-full justify-between'
                      : 'min-w-max justify-start'
                  }`}
                >
                  {processedData.length === 0 ? (
                    <div className="w-full flex items-center justify-center h-full text-xs text-gray-400">
                      No hay datos disponibles para mostrar con los filtros seleccionados.
                    </div>
                  ) : (
                    processedData.map((item) => {
                      const isSelected = selectedEntity === item.entidad;
                      const isHovered = hoveredEntity === item.entidad;

                      const valPrimary = metric === 'completo' ? item.porcCompleto : item.avance;
                      const valSecondary = item.porcCompleto;

                      const stylesPrimary = getSemaforoStyles(valPrimary);
                      const stylesSecondary = getSemaforoStyles(valSecondary);

                      // Calculate column heights (clamped 0 to 100)
                      const primaryHeightPct = Math.min(100, Math.max(0, valPrimary));
                      const secondaryHeightPct = Math.min(100, Math.max(0, valSecondary));

                      // Width calculations
                      const colWidthClass = isFitMode
                        ? 'flex-1 min-w-[20px] max-w-[54px]'
                        : metric === 'compare'
                        ? 'w-14 sm:w-16'
                        : 'w-10 sm:w-12';

                      return (
                        <div
                          key={item.id}
                          onMouseEnter={() => setHoveredEntity(item.entidad)}
                          onMouseLeave={() => setHoveredEntity(null)}
                          onClick={() => onSelectEntity && onSelectEntity(isSelected ? '' : item.entidad)}
                          className={`group flex flex-col items-center h-full justify-end cursor-pointer transition-all duration-150 ${colWidthClass} ${
                            hoveredEntity && !isHovered ? 'opacity-55' : 'opacity-100'
                          }`}
                        >
                          {/* Top Percentage Label */}
                          <div
                            className={`mb-1.5 text-[10px] sm:text-[11px] font-bold tabular-nums text-center transition-transform group-hover:scale-110 ${
                              isSelected
                                ? 'text-[#691C32] scale-105'
                                : 'text-gray-800'
                            }`}
                          >
                            {formatPercent(valPrimary)}
                          </div>

                          {/* Column Bar Area */}
                          <div className="w-full flex items-end justify-center h-[240px] relative">
                            {metric !== 'compare' ? (
                              /* Single Vertical Column */
                              <div className="w-full h-full flex items-end justify-center">
                                <div
                                  className={`w-full max-w-[38px] rounded-t-md transition-all duration-300 shadow-2xs relative ${
                                    isSelected
                                      ? 'ring-2 ring-[#691C32] ring-offset-1'
                                      : 'group-hover:brightness-110'
                                  }`}
                                  style={{
                                    height: `${primaryHeightPct}%`,
                                    backgroundColor: stylesPrimary.hexColor,
                                    minHeight: primaryHeightPct > 0 ? '4px' : '0px',
                                  }}
                                >
                                  {/* Subtle top cap highlight */}
                                  <div className="absolute top-0 inset-x-0 h-1 bg-white/25 rounded-t-md" />
                                </div>
                              </div>
                            ) : (
                              /* Comparative Double Columns */
                              <div className="w-full h-full flex items-end justify-center gap-1">
                                {/* Avance Column */}
                                <div
                                  className="w-1/2 rounded-t-sm transition-all duration-300 shadow-2xs"
                                  style={{
                                    height: `${primaryHeightPct}%`,
                                    backgroundColor: stylesPrimary.hexColor,
                                    minHeight: primaryHeightPct > 0 ? '4px' : '0px',
                                  }}
                                  title={`Avance: ${formatPercent(item.avance)}`}
                                />
                                {/* Completo Column */}
                                <div
                                  className="w-1/2 rounded-t-sm transition-all duration-300 bg-[#A68D5D] shadow-2xs"
                                  style={{
                                    height: `${secondaryHeightPct}%`,
                                    minHeight: secondaryHeightPct > 0 ? '4px' : '0px',
                                  }}
                                  title={`Completo: ${formatPercent(item.porcCompleto)}`}
                                />
                              </div>
                            )}
                          </div>

                          {/* Entity Label in X-Axis: Angled Rotation for High Legibility */}
                          <div className="w-full pt-2 flex flex-col items-center">
                            <div
                              className="text-[10px] sm:text-[11px] font-semibold text-gray-700 group-hover:text-[#691C32] whitespace-nowrap text-right transform -rotate-45 origin-top-right transition-colors"
                              style={{
                                width: '90px',
                                marginTop: '4px',
                                textOverflow: 'ellipsis',
                                overflow: 'hidden',
                              }}
                              title={item.entidad}
                            >
                              {item.entidad}
                            </div>
                            
                            {/* CLUES count summary tag below */}
                            <div className="text-[9px] text-gray-400 tabular-nums mt-10 whitespace-nowrap">
                              {item.cluesConInventario}/{item.metaClues}
                            </div>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 2. HORIZONTAL BAR CHART (Alternative View when user toggles) */}
      {/* ========================================================================= */}
      {orientation === 'horizontal' && (
        <div className="space-y-2 max-h-[540px] overflow-y-auto pr-1">
          {processedData.length === 0 ? (
            <div className="py-8 text-center text-xs text-gray-400">
              No hay entidades que coincidan con los filtros aplicados.
            </div>
          ) : (
            processedData.map((item) => {
              const styles = getSemaforoStyles(item.avance);
              const isSelected = selectedEntity === item.entidad;

              return (
                <div
                  key={item.id}
                  onClick={() => onSelectEntity && onSelectEntity(isSelected ? '' : item.entidad)}
                  className={`group flex items-center gap-2 sm:gap-4 py-1.5 px-2 rounded cursor-pointer transition-all ${
                    isSelected
                      ? 'bg-[#691C32]/10 border border-[#691C32]/30 ring-1 ring-[#691C32]/20'
                      : 'hover:bg-gray-50'
                  }`}
                  title={`Click para aislar ${item.entidad}: ${item.cluesConInventario} de ${item.metaClues} CLUES`}
                >
                  {/* Entity Name & CLUES count */}
                  <div className="w-36 sm:w-48 shrink-0">
                    <div className="text-xs font-semibold text-gray-800 group-hover:text-[#691C32] truncate transition-colors">
                      {item.entidad}
                    </div>
                    <div className="text-[10px] text-gray-400 tabular-nums">
                      {formatNumber(item.cluesConInventario)} / {formatNumber(item.metaClues)} CLUES
                    </div>
                  </div>

                  {/* Horizontal Progress Bar Track */}
                  <div className="flex-1 relative bg-gray-100 rounded-full h-4 overflow-hidden border border-gray-200/60">
                    {/* 70% and 90% Guidelines */}
                    <div
                      className="absolute top-0 bottom-0 left-[70%] w-px bg-amber-500 z-10"
                      title="Umbral 70%"
                    />
                    <div
                      className="absolute top-0 bottom-0 left-[90%] w-px bg-emerald-600 z-10"
                      title="Meta 90%"
                    />

                    {/* Filled Bar */}
                    <div
                      className="h-full transition-all duration-300 rounded-full"
                      style={{
                        width: `${Math.min(100, Math.max(2, item.avance))}%`,
                        backgroundColor: styles.hexColor,
                      }}
                    />
                  </div>

                  {/* Percentage & Semáforo Dot */}
                  <div className="w-20 sm:w-24 shrink-0 text-right flex items-center justify-end gap-1.5">
                    <span
                      className="w-2 h-2 rounded-full shrink-0"
                      style={{ backgroundColor: styles.hexColor }}
                      aria-hidden="true"
                    />
                    <span className="text-xs font-bold tabular-nums text-gray-800">
                      {formatPercent(item.avance)}
                    </span>
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}

      {/* Chart Footer with Institutional Reference Legend & Averages */}
      <div className="mt-4 pt-3 border-t border-gray-100 flex flex-col md:flex-row md:items-center justify-between text-xs text-gray-600 gap-3">
        {/* Color Legend Matching the Official Report */}
        <div className="flex flex-wrap items-center gap-3 text-[11px]">
          <span className="font-semibold text-gray-700">Semáforo:</span>
          <span className="inline-flex items-center gap-1">
            <span className="w-3 h-3 rounded-xs bg-[#137537]" />
            <span>≥ 70% Óptimo / Aceptable</span>
          </span>
          <span className="inline-flex items-center gap-1">
            <span className="w-3 h-3 rounded-xs bg-[#E5B824]" />
            <span>50% - 69.9% Regular</span>
          </span>
          <span className="inline-flex items-center gap-1">
            <span className="w-3 h-3 rounded-xs bg-[#C82333]" />
            <span>&lt; 50% Atención Prioritaria</span>
          </span>
          {metric === 'compare' && (
            <span className="inline-flex items-center gap-1 border-l border-gray-300 pl-2">
              <span className="w-3 h-3 rounded-xs bg-[#A68D5D]" />
              <span>% Inventario Completo</span>
            </span>
          )}
        </div>

        {/* Dynamic Summary Note */}
        <div className="flex items-center gap-2 text-[11px] text-gray-500">
          <Info className="w-3.5 h-3.5 text-[#691C32]" />
          <span>
            Mostrando <strong>{processedData.length}</strong> entidades · Promedio visible:{' '}
            <strong className="text-[#10312B]">{formatPercent(averageAvance)}</strong>
          </span>
        </div>
      </div>
    </section>
  );
};
