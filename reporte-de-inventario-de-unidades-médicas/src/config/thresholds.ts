import { SemaforoLevel } from '../types';

/**
 * Rangos de color para las celdas del inventario por entidad federativa.
 * Mantienen la misma lógica que la función semaforo del reporte HTML.
 */
export const SEMAFORO_THRESHOLDS = {
  OPTIMO: 90.0, // Meta óptima institucional
  ALTO: 70.0,   // >= 70% (Verde en reporte oficial)
  MEDIO: 50.0,  // 50% - 69.99% (Amarillo en reporte oficial)
  // < 50% (Rojo)
} as const;

export const INSTITUTIONAL_COLORS = {
  guinda: '#691C32',
  guindaHover: '#521426',
  guindaLight: '#FBF4F6',
  guindaBorder: '#E7C8D2',
  dorado: '#D4C19C',
  doradoDark: '#A68D5D',
  doradoLight: '#FAF7F0',
  verdeOscuro: '#10312B',
  verdeOscuroLight: '#E8EFEF',
  // Colores semáforo oficiales del reporte
  verdeOficial: '#137537',
  amarilloOficial: '#E5B824',
  rojoOficial: '#C82333',
} as const;

export function getSemaforoLevel(percentage: number): SemaforoLevel {
  if (percentage < 50) {
    return 'bajo';
  }
  if (percentage < 75) {
    return 'medio';
  }
  if (percentage < 100) {
    return 'alto';
  }
  return 'excelente';
}

function getTextColor(backgroundColor: string): string {
  return backgroundColor === '#D41111' || backgroundColor === '#0D5D2A'
    ? '#FFFFFF'
    : '#000000';
}

export function getSemaforoStyles(percentage: number): {
  dotColor: string;
  textColor: string;
  bgColor: string;
  borderColor: string;
  barColor: string;
  hexColor: string;
  label: string;
} {
  const level = getSemaforoLevel(percentage);
  switch (level) {
    case 'bajo':
      return {
        dotColor: 'bg-[#D41111]',
        textColor: 'text-[#D41111]',
        bgColor: 'bg-rose-50',
        borderColor: 'border-rose-300',
        barColor: 'bg-[#D41111]',
        hexColor: '#D41111',
        label: 'Atención prioritaria (<50%)',
      };
    case 'medio':
      return {
        dotColor: 'bg-[#F1D54A]',
        textColor: 'text-[#A67C00]',
        bgColor: 'bg-yellow-50',
        borderColor: 'border-yellow-300',
        barColor: 'bg-[#F1D54A]',
        hexColor: '#F1D54A',
        label: 'Seguimiento (50% - 74.9%)',
      };
    case 'alto':
      return {
        dotColor: 'bg-[#88A91E]',
        textColor: 'text-[#88A91E]',
        bgColor: 'bg-lime-50',
        borderColor: 'border-lime-300',
        barColor: 'bg-[#88A91E]',
        hexColor: '#88A91E',
        label: 'Avance aceptable (75% en adelante)',
      };
    case 'excelente':
    default:
      return {
        dotColor: 'bg-[#0D5D2A]',
        textColor: 'text-[#0D5D2A]',
        bgColor: 'bg-green-50',
        borderColor: 'border-green-300',
        barColor: 'bg-[#0D5D2A]',
        hexColor: '#0D5D2A',
        label: 'Avance completo (100% o más)',
      };
  }
}

/**
 * Estilos de celda sólida institucional (bloques verde, amarillo y rojo como en el reporte impreso/oficial)
 */
export function getSolidCellStyles(percentage: number): {
  bgClass: string;
  textClass: string;
  styleObj: React.CSSProperties;
} {
  const level = getSemaforoLevel(percentage);
  switch (level) {
    case 'bajo':
      return {
        bgClass: 'bg-[#D41111] text-white',
        textClass: 'text-white',
        styleObj: { backgroundColor: '#D41111', color: '#FFFFFF' },
      };
    case 'medio':
      return {
        bgClass: 'bg-[#F1D54A] text-gray-900',
        textClass: 'text-gray-900',
        styleObj: { backgroundColor: '#F1D54A', color: '#111827' },
      };
    case 'alto':
      return {
        bgClass: 'bg-[#88A91E] text-black',
        textClass: 'text-black',
        styleObj: { backgroundColor: '#88A91E', color: getTextColor('#88A91E') },
      };
    case 'excelente':
    default:
      return {
        bgClass: 'bg-[#0D5D2A] text-white',
        textClass: 'text-white',
        styleObj: { backgroundColor: '#0D5D2A', color: getTextColor('#0D5D2A') },
      };
  }
}
