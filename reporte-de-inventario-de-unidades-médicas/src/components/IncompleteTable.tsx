import React, { useState, useMemo } from 'react';
import { UnitDetail, SortDirection } from '../types';
import { AlertTriangle, ArrowUpDown, ArrowUp, ArrowDown, Search, MapPin, Building2 } from 'lucide-react';

interface IncompleteTableProps {
  units: UnitDetail[];
  selectedEntity: string;
  onSelectEntity: (entity: string) => void;
}

type IncompleteSortField = 'clues' | 'nombreUnidad' | 'entidad' | 'motivoIncompleto';

export const IncompleteTable: React.FC<IncompleteTableProps> = ({
  units,
  selectedEntity,
  onSelectEntity,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [entityFilter, setEntityFilter] = useState(selectedEntity);
  const [sortField, setSortField] = useState<IncompleteSortField>('entidad');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  React.useEffect(() => {
    setEntityFilter(selectedEntity);
  }, [selectedEntity]);

  const availableEntities = useMemo(() => {
    const set = new Set<string>();
    units.forEach((u) => set.add(u.entidad));
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
  }, [units]);

  const filteredUnits = useMemo(() => {
    return units
      .filter((u) => {
        if (entityFilter && u.entidad !== entityFilter) return false;
        if (searchTerm.trim() !== '') {
          const term = searchTerm.toLowerCase();
          const matchClues = u.clues.toLowerCase().includes(term);
          const matchName = u.nombreUnidad.toLowerCase().includes(term);
          const matchEntity = u.entidad.toLowerCase().includes(term);
          const matchMotivo = (u.motivoIncompleto || '').toLowerCase().includes(term);
          if (!matchClues && !matchName && !matchEntity && !matchMotivo) return false;
        }
        return true;
      })
      .sort((a, b) => {
        let valA = '';
        let valB = '';

        switch (sortField) {
          case 'clues':
            valA = a.clues;
            valB = b.clues;
            break;
          case 'nombreUnidad':
            valA = a.nombreUnidad;
            valB = b.nombreUnidad;
            break;
          case 'entidad':
            valA = a.entidad;
            valB = b.entidad;
            break;
          case 'motivoIncompleto':
            valA = a.motivoIncompleto || '';
            valB = b.motivoIncompleto || '';
            break;
        }

        return sortDirection === 'asc'
          ? valA.localeCompare(valB, 'es')
          : valB.localeCompare(valA, 'es');
      });
  }, [units, entityFilter, searchTerm, sortField, sortDirection]);

  const handleSort = (field: IncompleteSortField) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection('asc');
    }
  };

  const renderSortIcon = (field: IncompleteSortField) => {
    if (sortField !== field) {
      return <ArrowUpDown className="w-3 h-3 text-gray-300 opacity-60" />;
    }
    return sortDirection === 'asc' ? (
      <ArrowUp className="w-3 h-3 text-[#691C32]" />
    ) : (
      <ArrowDown className="w-3 h-3 text-[#691C32]" />
    );
  };

  return (
    <section
      id="section-clues-incompletos"
      aria-labelledby="heading-clues-incompletos"
      className="bg-white rounded-lg border border-gray-200 shadow-xs overflow-hidden print:border-gray-300 print:shadow-none"
    >
      {/* Table Header */}
      <div className="p-4 sm:p-5 border-b border-gray-100 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <div className="p-1 rounded bg-amber-100 text-amber-800">
              <AlertTriangle className="w-4 h-4" />
            </div>
            <h2
              id="heading-clues-incompletos"
              className="text-base sm:text-lg font-bold text-[#10312B] font-['Montserrat',sans-serif]"
            >
              CLUES incompletos ({filteredUnits.length})
            </h2>
          </div>
          <p className="text-xs text-gray-500 mt-1">
            Unidades que han reportado solo un catálogo (medicamentos o curación) o presentan registro inconcluso.
          </p>
        </div>

        {/* Local Search & Filter */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 text-xs print:hidden">
          {/* Entity Filter */}
          <div className="relative min-w-[160px]">
            <div className="absolute inset-y-0 left-0 pl-2 flex items-center pointer-events-none text-gray-400">
              <MapPin className="w-3 h-3" />
            </div>
            <select
              value={entityFilter}
              onChange={(e) => {
                setEntityFilter(e.target.value);
                onSelectEntity(e.target.value);
              }}
              className="w-full pl-7 pr-6 py-1 bg-gray-50 border border-gray-200 rounded text-xs text-gray-700 focus:ring-1 focus:ring-[#691C32]"
            >
              <option value="">Todas las entidades</option>
              {availableEntities.map((ent) => (
                <option key={ent} value={ent}>
                  {ent}
                </option>
              ))}
            </select>
          </div>

          {/* Search box */}
          <div className="relative min-w-[180px]">
            <div className="absolute inset-y-0 left-0 pl-2 flex items-center pointer-events-none text-gray-400">
              <Search className="w-3 h-3" />
            </div>
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Buscar CLUES o unidad..."
              className="w-full pl-7 pr-2 py-1 bg-gray-50 border border-gray-200 rounded text-xs text-gray-700 focus:ring-1 focus:ring-[#691C32]"
            />
          </div>
        </div>
      </div>

      {/* Table Content */}
      <div className="overflow-x-auto">
        <table className="w-full text-xs text-left border-collapse">
          <thead className="bg-gray-50 border-b border-gray-200 text-gray-700 select-none">
            <tr>
              <th
                scope="col"
                onClick={() => handleSort('clues')}
                className="py-2.5 px-3.5 font-bold uppercase tracking-wider text-[11px] text-[#10312B] cursor-pointer group hover:bg-gray-100 transition-colors w-36"
              >
                <div className="flex items-center gap-1.5">
                  <span>CLUES IMB</span>
                  {renderSortIcon('clues')}
                </div>
              </th>

              <th
                scope="col"
                onClick={() => handleSort('nombreUnidad')}
                className="py-2.5 px-3.5 font-bold uppercase tracking-wider text-[11px] text-gray-700 cursor-pointer group hover:bg-gray-100 transition-colors"
              >
                <div className="flex items-center gap-1.5">
                  <span>Nombre de la unidad</span>
                  {renderSortIcon('nombreUnidad')}
                </div>
              </th>

              <th
                scope="col"
                onClick={() => handleSort('entidad')}
                className="py-2.5 px-3.5 font-bold uppercase tracking-wider text-[11px] text-gray-700 cursor-pointer group hover:bg-gray-100 transition-colors w-44"
              >
                <div className="flex items-center gap-1.5">
                  <span>Entidad</span>
                  {renderSortIcon('entidad')}
                </div>
              </th>

              <th
                scope="col"
                onClick={() => handleSort('motivoIncompleto')}
                className="py-2.5 px-3.5 font-bold uppercase tracking-wider text-[11px] text-gray-700 cursor-pointer group hover:bg-gray-100 transition-colors"
              >
                <div className="flex items-center gap-1.5">
                  <span>Conteo / Estatus Faltante</span>
                  {renderSortIcon('motivoIncompleto')}
                </div>
              </th>
            </tr>
          </thead>

          <tbody className="divide-y divide-gray-100">
            {filteredUnits.length === 0 ? (
              <tr>
                <td colSpan={4} className="py-8 text-center text-gray-400">
                  No hay unidades incompletas que coincidan con los criterios.
                </td>
              </tr>
            ) : (
              filteredUnits.map((unit, idx) => (
                <tr
                  key={unit.id}
                  className={`transition-colors ${
                    idx % 2 === 0 ? 'bg-white hover:bg-amber-50/30' : 'bg-gray-50/40 hover:bg-amber-50/30'
                  }`}
                >
                  <td className="py-2.5 px-3.5 font-mono font-semibold text-gray-900 whitespace-nowrap">
                    {unit.clues}
                  </td>
                  <td className="py-2.5 px-3.5 text-gray-800">
                    <div className="font-medium text-gray-900">{unit.nombreUnidad}</div>
                    <div className="text-[11px] text-gray-400 flex items-center gap-1.5 mt-0.5">
                      <Building2 className="w-3 h-3 text-gray-400" />
                      <span>{unit.tipoUnidad}</span>
                      {unit.municipio && <span>· Mun. {unit.municipio}</span>}
                      {unit.ultimaFechaRegistro && <span>· Última carga: {unit.ultimaFechaRegistro}</span>}
                    </div>
                  </td>
                  <td className="py-2.5 px-3.5 text-gray-700 whitespace-nowrap">
                    <span className="inline-block px-2 py-0.5 bg-gray-100 rounded text-gray-700 text-[11px] font-medium">
                      {unit.entidad}
                    </span>
                  </td>
                  <td className="py-2.5 px-3.5">
                    <div className="flex items-center gap-2">
                      <span className="inline-flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1 rounded bg-amber-50 text-amber-900 border border-amber-200">
                        <span className="w-1.5 h-1.5 rounded-full bg-amber-500" />
                        {unit.motivoIncompleto || 'Inventario parcial'}
                      </span>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
};
