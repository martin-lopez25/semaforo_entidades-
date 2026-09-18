export interface EntityInventory {
  id: string;
  entidad: string;
  metaClues: number;
  cluesConInventario: number;
  cluesMedicamentos: number; // 010 / 040
  cluesMaterialCuracion: number; // 060
  inventarioCompleto: number;
}

export interface UnitDetail {
  id: string;
  clues: string;
  nombreUnidad: string;
  entidad: string;
  municipio?: string;
  tipoUnidad: string; // 'Primer Nivel' | 'Segundo Nivel' | 'Hospital Comunitario'
  diasSinReporte?: number;
  motivoIncompleto?: string; // 'Falta Material 060' | 'Falta Medicamentos 010/040' | 'Parcial ambos'
  avanceMedicamentos?: number;
  avanceMaterial?: number;
  ultimaFechaRegistro?: string;
}

export interface NationalSummary {
  metaTotal: number;
  conInventarioTotal: number;
  porcentajeAvance: number;
  medicamentosTotal: number;
  materialCuracionTotal: number;
  inventarioCompletoTotal: number;
  porcentajeCompleto: number;
}

export type SemaforoLevel = 'excelente' | 'alto' | 'medio' | 'bajo';

export type CaptureTarget =
  | 'chart'
  | 'summary'
  | 'inventory'
  | 'not-reported'
  | 'incomplete'
  | 'full';

export type SortDirection = 'asc' | 'desc';

export type EntitySortField =
  | 'entidad'
  | 'metaClues'
  | 'cluesConInventario'
  | 'porcentajeAvance'
  | 'cluesMedicamentos'
  | 'cluesMaterialCuracion'
  | 'porcentajeCompleto'
  | 'inventarioCompleto';
