import React, { useState, useMemo } from 'react';
import { EntityInventory, EntitySortField, SortDirection } from '../types';
import { calculateAvance, calculateCompleto } from '../utils/calculations';
import { formatNumber, formatPercent } from '../utils/formatters';
import { getSolidCellStyles, getSemaforoStyles } from '../config/thresholds';
import { StatusBadge } from './StatusBadge';
import {
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  HelpCircle,
  Layers,
  Palette,
  CheckSquare,
} from 'lucide-react';

interface InventoryTableProps {
  data: EntityInventory[];
  selectedEntity: string;
  onSelectEntity: (entity: string) => void;
  notReportedCount?: number;
}

export const InventoryTable: React.FC<InventoryTableProps> = ({
  data,
  selectedEntity,
  onSelectEntity,
  notReportedCount = 0,
}) => {
  const [sortField, setSortField] = useState<EntitySortField>('porcentajeAvance');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');
  // Official solid blocks style (matching reference screenshot) is default
  const [useSolidCellColors, setUseSolidCellColors] = useState<boolean>(true);

  // Compute sorting
  const sortedData = useMemo(() => {
    const list = [...data];
    list.sort((a, b) => {
      let valA: number | string = 0;
      let valB: number | string = 0;

      switch (sortField) {
        case 'entidad':
          valA = a.entidad;
          valB = b.entidad;
          return sortDirection === 'asc'
            ? valA.localeCompare(valB, 'es')
            : valB.localeCompare(valA, 'es');
        case 'metaClues':
          valA = a.metaClues;
          valB = b.metaClues;
          break;
        case 'cluesConInventario':
          valA = a.cluesConInventario;
          valB = b.cluesConInventario;
          break;
        case 'porcentajeAvance':
          valA = calculateAvance(a.cluesConInventario, a.metaClues);
          valB = calculateAvance(b.cluesConInventario, b.metaClues);
          break;
        case 'cluesMedicamentos':
          valA = a.cluesMedicamentos;
          valB = b.cluesMedicamentos;
          break;
        case 'cluesMaterialCuracion':
          valA = a.cluesMaterialCuracion;
          valB = b.cluesMaterialCuracion;
          break;
        case 'porcentajeCompleto':
          valA = calculateCompleto(a.inventarioCompleto, a.metaClues);
          valB = calculateCompleto(b.inventarioCompleto, b.metaClues);
          break;
        case 'inventarioCompleto':
          valA = a.inventarioCompleto;
          valB = b.inventarioCompleto;
          break;
      }

      if (typeof valA === 'number' && typeof valB === 'number') {
        return sortDirection === 'asc' ? valA - valB : valB - valA;
      }
      return 0;
    });

    return list;
  }, [data, sortField, sortDirection]);

  // Aggregate totals
  const totals = useMemo(() => {
    const meta = data.reduce((sum, item) => sum + item.metaClues, 0);
    const conInventario = data.reduce((sum, item) => sum + item.cluesConInventario, 0);
    const medicamentos = data.reduce((sum, item) => sum + item.cluesMedicamentos, 0);
    const material = data.reduce((sum, item) => sum + item.cluesMaterialCuracion, 0);
    const completo = data.reduce((sum, item) => sum + item.inventarioCompleto, 0);
    const avance = calculateAvance(conInventario, meta);
    const porcCompleto = calculateCompleto(completo, meta);

    return {
      meta,
      conInventario,
      medicamentos,
      material,
      completo,
      avance,
      porcCompleto,
    };
  }, [data]);

  const handleSort = (field: EntitySortField) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection(field === 'entidad' ? 'asc' : 'desc');
    }
  };

  const renderSortIcon = (field: EntitySortField) => {
    if (sortField !== field) {
      return <ArrowUpDown className="w-3 h-3 text-white/40 opacity-70 group-hover:opacity-100" />;
    }
    return sortDirection === 'asc' ? (
      <ArrowUp className="w-3 h-3 text-[#D4C19C]" />
    ) : (
      <ArrowDown className="w-3 h-3 text-[#D4C19C]" />
    );
  };

  return (
    <section
      id="capture-inventory"
      aria-labelledby="heading-tabla-entidades"
      className="bg-white rounded-lg border border-gray-200 shadow-xs overflow-visible print:border-gray-300 print:shadow-none"
    >
      {/* Table Section Header */}
      <div className="p-4 sm:p-5 border-b border-gray-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <Layers className="w-5 h-5 text-[#691C32]" />
            <h2
              id="heading-tabla-entidades"
              className="text-base sm:text-lg font-bold text-[#10312B] font-['Montserrat',sans-serif]"
            >
              Inventario por entidad federativa
            </h2>
          </div>
          <p className="text-xs text-gray-500 mt-0.5">
            Desglose analítico oficial con meta de CLUES, avance y cobertura de catálogos
          </p>
        </div>

        <div className="flex items-center gap-3 print:hidden">
          {/* Style Toggle (Solid Cells like screenshot vs Badges) */}
          <button
            type="button"
            onClick={() => setUseSolidCellColors(!useSolidCellColors)}
            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded text-xs font-semibold border transition-all ${
              useSolidCellColors
                ? 'bg-[#10312B] text-white border-[#10312B] shadow-2xs'
                : 'bg-gray-50 text-gray-700 border-gray-200 hover:bg-gray-100'
            }`}
            title="Alternar entre celdas sólidas (formato oficial) e insignias"
          >
            <Palette className="w-3.5 h-3.5" />
            <span>{useSolidCellColors ? 'Celdas Oficiales' : 'Estilo Insignia'}</span>
          </button>

          <div className="text-xs text-gray-500 font-medium border-l border-gray-200 pl-3">
            Registros: <span className="font-bold text-[#10312B]">{sortedData.length}</span>
          </div>
        </div>
      </div>

      {/* Table with Sticky Official Deep Red/Guinda Header */}
      <div className="overflow-x-auto relative">
        <table className="w-full text-xs text-left border-collapse">
          {/* Table Head: Official Deep Guinda Background #691C32 */}
          <thead className="sticky top-0 z-20 bg-[#691C32] text-white shadow-xs select-none">
            <tr>
              {/* entidad */}
              <th
                scope="col"
                onClick={() => handleSort('entidad')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white cursor-pointer group hover:bg-[#571629] transition-colors"
              >
                <div className="flex items-center gap-1.5">
                  <span>entidad</span>
                  {renderSortIcon('entidad')}
                </div>
              </th>

              {/* meta de clues */}
              <th
                scope="col"
                onClick={() => handleSort('metaClues')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors"
              >
                <div className="flex items-center justify-end gap-1.5">
                  <span>meta de clues</span>
                  {renderSortIcon('metaClues')}
                </div>
              </th>

              {/* clues con inventario */}
              <th
                scope="col"
                onClick={() => handleSort('cluesConInventario')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors"
              >
                <div className="flex items-center justify-end gap-1.5">
                  <span>clues con inventario</span>
                  {renderSortIcon('cluesConInventario')}
                </div>
              </th>

              {/* pct avance */}
              <th
                scope="col"
                onClick={() => handleSort('porcentajeAvance')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors bg-[#571629]"
              >
                <div className="flex items-center justify-end gap-1.5">
                  <span>pct avance</span>
                  {renderSortIcon('porcentajeAvance')}
                </div>
              </th>

              {/* clues medicamentos 010 040 */}
              <th
                scope="col"
                onClick={() => handleSort('cluesMedicamentos')}
                className="py-3 px-3 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors"
                title="Claves de Medicamentos (Catálogos 010 y 040)"
              >
                <div className="flex items-center justify-end gap-1">
                  <span>clues medicamentos 010 040</span>
                  {renderSortIcon('cluesMedicamentos')}
                </div>
              </th>

              {/* clues material curacion 060 */}
              <th
                scope="col"
                onClick={() => handleSort('cluesMaterialCuracion')}
                className="py-3 px-3 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors"
                title="Claves de Material de Curación (Catálogo 060)"
              >
                <div className="flex items-center justify-end gap-1">
                  <span>clues material curacion 060</span>
                  {renderSortIcon('cluesMaterialCuracion')}
                </div>
              </th>

              {/* pct completo */}
              <th
                scope="col"
                onClick={() => handleSort('porcentajeCompleto')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors bg-[#571629]"
              >
                <div className="flex items-center justify-end gap-1.5">
                  <span>pct completo</span>
                  {renderSortIcon('porcentajeCompleto')}
                </div>
              </th>

              {/* inventario completo */}
              <th
                scope="col"
                onClick={() => handleSort('inventarioCompleto')}
                className="py-3 px-3.5 font-bold uppercase tracking-wider text-[11px] text-white text-right cursor-pointer group hover:bg-[#571629] transition-colors"
              >
                <div className="flex items-center justify-end gap-1.5">
                  <span>inventario completo</span>
                  {renderSortIcon('inventarioCompleto')}
                </div>
              </th>
            </tr>
          </thead>

          {/* Table Body */}
          <tbody className="divide-y divide-gray-200 font-normal">
            {sortedData.length === 0 ? (
              <tr>
                <td colSpan={8} className="py-8 text-center text-gray-400">
                  No se encontraron resultados con los filtros actuales.
                </td>
              </tr>
            ) : (
              sortedData.map((item, idx) => {
                const avance = calculateAvance(item.cluesConInventario, item.metaClues);
                const porcCompleto = calculateCompleto(item.inventarioCompleto, item.metaClues);
                const isSelected = selectedEntity === item.entidad;

                // Solid cell styling matching the official report
                const avanceSolid = getSolidCellStyles(avance);
                const completoSolid = getSolidCellStyles(porcCompleto);

                return (
                  <tr
                    key={item.id}
                    onClick={() => onSelectEntity(isSelected ? '' : item.entidad)}
                    className={`transition-colors cursor-pointer group ${
                      isSelected
                        ? 'bg-[#691C32]/10 font-semibold ring-1 ring-inset ring-[#691C32]/30'
                        : idx % 2 === 0
                        ? 'bg-white hover:bg-amber-50/40'
                        : 'bg-gray-50/50 hover:bg-amber-50/40'
                    }`}
                  >
                    {/* entidad */}
                    <td className="py-2 px-3.5 text-gray-900 font-medium whitespace-nowrap group-hover:text-[#691C32] uppercase">
                      <div className="flex items-center gap-2">
                        {isSelected && (
                          <span className="w-1.5 h-3.5 bg-[#691C32] rounded-full" />
                        )}
                        <span>{item.entidad}</span>
                      </div>
                    </td>

                    {/* meta de clues */}
                    <td className="py-2 px-3.5 text-right font-medium text-gray-700 tabular-nums whitespace-nowrap">
                      {item.metaClues.toFixed(1)}
                    </td>

                    {/* clues con inventario */}
                    <td className="py-2 px-3.5 text-right font-semibold text-gray-900 tabular-nums whitespace-nowrap">
                      {item.cluesConInventario.toFixed(1)}
                    </td>

                    {/* pct avance (Solid Block in Official Mode) */}
                    {useSolidCellColors ? (
                      <td
                        className={`py-2 px-3.5 text-center font-bold tabular-nums whitespace-nowrap ${avanceSolid.bgClass}`}
                        style={avanceSolid.styleObj}
                      >
                        {avance.toFixed(2)}
                      </td>
                    ) : (
                      <td className="py-2 px-3.5 text-right tabular-nums whitespace-nowrap bg-gray-50/50">
                        <StatusBadge percentage={avance} size="sm" />
                      </td>
                    )}

                    {/* clues medicamentos 010 040 */}
                    <td className="py-2 px-3 text-right text-gray-700 tabular-nums whitespace-nowrap">
                      {item.cluesMedicamentos.toFixed(1)}
                    </td>

                    {/* clues material curacion 060 */}
                    <td className="py-2 px-3 text-right text-gray-700 tabular-nums whitespace-nowrap">
                      {item.cluesMaterialCuracion.toFixed(1)}
                    </td>

                    {/* pct completo */}
                    <td className="py-2 px-3.5 text-right font-semibold text-gray-800 tabular-nums whitespace-nowrap bg-gray-50/30">
                      {porcCompleto.toFixed(1)}
                    </td>

                    {/* inventario completo (Solid Block in Official Mode) */}
                    {useSolidCellColors ? (
                      <td
                        className={`py-2 px-3.5 text-center font-bold tabular-nums whitespace-nowrap ${completoSolid.bgClass}`}
                        style={completoSolid.styleObj}
                      >
                        {porcCompleto.toFixed(2)}
                      </td>
                    ) : (
                      <td className="py-2 px-3.5 text-right font-bold text-[#10312B] tabular-nums whitespace-nowrap">
                        <StatusBadge percentage={porcCompleto} size="sm" />
                      </td>
                    )}
                  </tr>
                );
              })
            )}
          </tbody>

          {/* Table Footer: Totales Nacionales */}
          <tfoot className="sticky bottom-0 z-10 bg-[#10312B] text-white font-semibold text-xs border-t-2 border-[#D4C19C]">
            <tr>
              <td className="py-3 px-3.5 uppercase tracking-wider font-bold">
                Totales ({sortedData.length} Entidades)
              </td>
              <td className="py-3 px-3.5 text-right tabular-nums font-bold">
                {totals.meta.toFixed(1)}
              </td>
              <td className="py-3 px-3.5 text-right tabular-nums font-bold text-[#D4C19C]">
                {totals.conInventario.toFixed(1)}
              </td>
              <td className="py-3 px-3.5 text-center tabular-nums font-bold bg-[#137537] text-white">
                {totals.avance.toFixed(2)}
              </td>
              <td className="py-3 px-3 text-right tabular-nums text-white/90">
                {totals.medicamentos.toFixed(1)}
              </td>
              <td className="py-3 px-3 text-right tabular-nums text-white/90">
                {totals.material.toFixed(1)}
              </td>
              <td className="py-3 px-3.5 text-right tabular-nums font-bold text-[#D4C19C]">
                {totals.porcCompleto.toFixed(1)}
              </td>
              <td className="py-3 px-3.5 text-center tabular-nums font-bold bg-[#137537] text-white">
                {totals.porcCompleto.toFixed(2)}
              </td>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* Table Footer Guide */}
      <div className="px-4 py-2.5 bg-gray-50 border-t border-gray-200 text-gray-500 flex flex-col sm:flex-row sm:items-center justify-between gap-1.5">
        <span className="text-[11px]">
          CLUES que NO reportaron: <strong className="text-[#691C32] text-base font-bold">{notReportedCount}</strong>
        </span>
        <span className="text-gray-400 text-[11px]">
          Medicamentos: Catálogo 010/040 · Material de Curación: Catálogo 060
        </span>
      </div>
    </section>
  );
};
