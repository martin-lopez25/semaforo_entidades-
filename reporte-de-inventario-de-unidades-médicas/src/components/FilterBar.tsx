import React from 'react';
import { Search, MapPin, X, Filter } from 'lucide-react';

interface FilterBarProps {
  selectedEntity: string;
  onSelectEntity: (entity: string) => void;
  searchTerm: string;
  onSearchChange: (term: string) => void;
  entitiesList: string[];
  totalMatchesCount?: number;
}

export const FilterBar: React.FC<FilterBarProps> = ({
  selectedEntity,
  onSelectEntity,
  searchTerm,
  onSearchChange,
  entitiesList,
  totalMatchesCount,
}) => {
  const hasActiveFilters = selectedEntity !== '' || searchTerm.trim() !== '';

  const handleResetFilters = () => {
    onSelectEntity('');
    onSearchChange('');
  };

  return (
    <div
      id="filter-bar"
      className="bg-white rounded-lg border border-gray-200 p-3 sm:p-4 shadow-xs print:hidden"
    >
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
        {/* Filters title / indicator */}
        <div className="flex flex-col gap-1.5">
          <div className="flex items-center gap-2 text-xs font-semibold text-[#10312B] uppercase tracking-wider">
            <Filter className="w-3.5 h-3.5 text-[#691C32]" />
            <span>Filtros de consulta</span>
            {hasActiveFilters && totalMatchesCount !== undefined && (
              <span className="ml-1 text-[11px] bg-[#691C32]/10 text-[#691C32] px-2 py-0.5 rounded-full font-medium">
                {totalMatchesCount} entidades visibles
              </span>
            )}
          </div>
          <span className="text-[11px] text-gray-500">
            Selecciona cualquier entidad para aislarla en el reporte. Haz clic en las cabeceras para ordenar.
          </span>
        </div>

        {/* Inputs */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 flex-1 max-w-2xl justify-end">
          {/* Entity Selector */}
          <div className="relative min-w-[220px]">
            <div className="absolute inset-y-0 left-0 pl-2.5 flex items-center pointer-events-none text-gray-400">
              <MapPin className="w-3.5 h-3.5 text-[#691C32]" />
            </div>
            <select
              id="select-entidad"
              value={selectedEntity}
              onChange={(e) => onSelectEntity(e.target.value)}
              className="w-full pl-8 pr-8 py-1.5 text-xs bg-gray-50 border border-gray-300 rounded focus:ring-1 focus:ring-[#691C32] focus:border-[#691C32] text-gray-800 transition-colors cursor-pointer"
            >
              <option value="">Todas las entidades ({entitiesList.length})</option>
              {entitiesList.map((entidad) => (
                <option key={entidad} value={entidad}>
                  {entidad}
                </option>
              ))}
            </select>
          </div>

          {/* Search Term */}
          <div className="relative flex-1 min-w-[200px]">
            <div className="absolute inset-y-0 left-0 pl-2.5 flex items-center pointer-events-none text-gray-400">
              <Search className="w-3.5 h-3.5" />
            </div>
            <input
              id="input-search"
              type="text"
              value={searchTerm}
              onChange={(e) => onSearchChange(e.target.value)}
              placeholder="Buscar entidad, CLUES o unidad..."
              className="w-full pl-8 pr-7 py-1.5 text-xs bg-gray-50 border border-gray-300 rounded focus:ring-1 focus:ring-[#691C32] focus:border-[#691C32] text-gray-800 placeholder-gray-400 transition-colors"
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => onSearchChange('')}
                className="absolute inset-y-0 right-0 pr-2 flex items-center text-gray-400 hover:text-gray-600"
                title="Limpiar búsqueda"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Reset Filters button */}
          {hasActiveFilters && (
            <button
              id="btn-reset-filters"
              type="button"
              onClick={handleResetFilters}
              className="inline-flex items-center justify-center gap-1 px-2.5 py-1.5 text-xs font-medium text-gray-600 hover:text-[#691C32] bg-gray-100 hover:bg-gray-200/80 rounded border border-gray-200 transition-colors shrink-0"
              title="Restablecer todos los filtros"
            >
              <X className="w-3.5 h-3.5" />
              <span>Limpiar filtros</span>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
