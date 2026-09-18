import { EntityInventory, NationalSummary } from '../types';

/**
 * Calcula el porcentaje de avance respecto a la meta.
 */
export function calculateAvance(conInventario: number, meta: number): number {
  if (!meta || meta <= 0) return 0;
  const val = (conInventario / meta) * 100;
  return Math.min(100, Math.max(0, val));
}

/**
 * Calcula el porcentaje de inventario completo respecto a la meta.
 */
export function calculateCompleto(completo: number, meta: number): number {
  if (!meta || meta <= 0) return 0;
  const val = (completo / meta) * 100;
  return Math.min(100, Math.max(0, val));
}

/**
 * Genera el resumen ejecutivo agregado a partir de un conjunto de entidades.
 */
export function calculateNationalSummary(entities: EntityInventory[]): NationalSummary {
  const metaTotal = entities.reduce((acc, curr) => acc + curr.metaClues, 0);
  const conInventarioTotal = entities.reduce((acc, curr) => acc + curr.cluesConInventario, 0);
  const medicamentosTotal = entities.reduce((acc, curr) => acc + curr.cluesMedicamentos, 0);
  const materialCuracionTotal = entities.reduce((acc, curr) => acc + curr.cluesMaterialCuracion, 0);
  const inventarioCompletoTotal = entities.reduce((acc, curr) => acc + curr.inventarioCompleto, 0);

  const porcentajeAvance = calculateAvance(conInventarioTotal, metaTotal);
  const porcentajeCompleto = calculateCompleto(inventarioCompletoTotal, metaTotal);

  return {
    metaTotal,
    conInventarioTotal,
    porcentajeAvance,
    medicamentosTotal,
    materialCuracionTotal,
    inventarioCompletoTotal,
    porcentajeCompleto,
  };
}
