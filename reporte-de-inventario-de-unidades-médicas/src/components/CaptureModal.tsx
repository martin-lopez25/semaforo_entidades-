import React, { useState, useEffect } from 'react';
import { toBlob, toPng } from 'html-to-image';
import { CaptureTarget } from '../types';
import { Download, MessageCircle, X, Check, Loader2, Image as ImageIcon } from 'lucide-react';

interface CaptureModalProps {
  isOpen: boolean;
  onClose: () => void;
  target: CaptureTarget;
  onChangeTarget: (target: CaptureTarget) => void;
}

export const CaptureModal: React.FC<CaptureModalProps> = ({
  isOpen,
  onClose,
  target,
  onChangeTarget,
}) => {
  const [loading, setLoading] = useState(false);
  const [imageDataUrl, setImageDataUrl] = useState<string | null>(null);
  const [imageBlob, setImageBlob] = useState<Blob | null>(null);
  const [shareSuccess, setShareSuccess] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  useEffect(() => {
    if (!isOpen) {
      setImageDataUrl(null);
      setImageBlob(null);
      setErrorMessage(null);
      return;
    }

    let isMounted = true;
    setLoading(true);
    setImageDataUrl(null);
    setImageBlob(null);
    setErrorMessage(null);

    // Give the DOM a moment to settle
    const timer = setTimeout(async () => {
      try {
        let elementId = 'capture-chart';
        if (target === 'summary') elementId = 'capture-summary';
        if (target === 'inventory') elementId = 'capture-inventory';
        if (target === 'not-reported') elementId = 'section-no-reportaron';
        if (target === 'incomplete') elementId = 'section-clues-incompletos';
        if (target === 'full') elementId = 'capture-full-report';

        const node = document.getElementById(elementId);
        if (!node) {
          throw new Error(`Elemento a capturar (${elementId}) no encontrado.`);
        }

        const dataUrl = await toPng(node, {
          backgroundColor: '#FFFFFF',
          pixelRatio: 2,
          quality: 0.98,
          cacheBust: true,
          width: node.scrollWidth || node.clientWidth,
          height: node.scrollHeight || node.clientHeight,
          style: {
            margin: '0',
            maxWidth: 'none',
            maxHeight: 'none',
            overflow: 'visible',
          },
        });
        const blob = await toBlob(node, {
          backgroundColor: '#FFFFFF',
          pixelRatio: 2,
          quality: 0.98,
          width: node.scrollWidth || node.clientWidth,
          height: node.scrollHeight || node.clientHeight,
        });

        if (isMounted) {
          setImageDataUrl(dataUrl);
          setImageBlob(blob);
          setLoading(false);
        }
      } catch (err: unknown) {
        console.error('Error al capturar imagen:', err);
        if (isMounted) {
          setErrorMessage(
            err instanceof Error ? err.message : 'No fue posible generar la imagen.'
          );
          setLoading(false);
        }
      }
    }, 200);

    return () => {
      isMounted = false;
      clearTimeout(timer);
    };
  }, [isOpen, target]);

  if (!isOpen) return null;

  const handleDownload = () => {
    if (!imageDataUrl) return;
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const link = document.createElement('a');
    link.download = `reporte-inventario-${target}-${dateStr}.png`;
    link.href = imageDataUrl;
    link.click();
  };

  const handleShare = async () => {
    if (!imageBlob) return;

    const fileName = `reporte-inventario-${target}.png`;
    const file = new File([imageBlob], fileName, { type: 'image/png' });

    if (navigator.share && navigator.canShare?.({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: getTargetTitle(),
          text: 'Reporte de Inventario de Unidades Médicas',
        });
        setShareSuccess(true);
        setTimeout(() => setShareSuccess(false), 3000);
      } catch (err: unknown) {
        if (!(err instanceof Error && err.name === 'AbortError')) {
          console.error('Error al compartir el PNG:', err);
        }
      }
      return;
    }

    handleDownload();
  };

  const getTargetTitle = () => {
    switch (target) {
      case 'chart':
        return 'Gráfica de Avance por Entidad';
      case 'summary':
        return 'Vista General (Métricas Ejecutivas)';
      case 'inventory':
        return 'Inventario por Entidad Federativa';
      case 'not-reported':
        return 'CLUES que no reportaron';
      case 'incomplete':
        return 'CLUES incompletos';
      case 'full':
        return 'Reporte Completo Institucional';
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 overflow-y-auto bg-black/60 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs animate-in fade-in duration-200"
      role="dialog"
      aria-modal="true"
      aria-labelledby="modal-capture-title"
    >
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full overflow-hidden border border-gray-200 flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-5 py-4 border-b border-gray-100 flex items-center justify-between bg-[#691C32] text-white">
          <div className="flex items-center gap-2">
            <ImageIcon className="w-5 h-5 text-[#D4C19C]" />
            <div>
              <h3 id="modal-capture-title" className="text-sm sm:text-base font-bold font-['Montserrat',sans-serif]">
                Captura de Reporte Institucional
              </h3>
              <p className="text-[11px] text-white/80">
                Generación de imagen PNG en alta resolución (2x) con fondo blanco
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1 rounded hover:bg-white/10 text-white/80 hover:text-white transition-colors"
            title="Cerrar modal"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Target Switcher */}
        <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="font-semibold text-gray-700">Área seleccionada:</div>
          <div className="inline-flex rounded-md border border-gray-300 p-0.5 bg-white">
            <button
              type="button"
              onClick={() => onChangeTarget('chart')}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                target === 'chart'
                  ? 'bg-[#10312B] text-white'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Gráfica
            </button>
            <button
              type="button"
              onClick={() => onChangeTarget('summary')}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                target === 'summary'
                  ? 'bg-[#691C32] text-white'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Vista General
            </button>
            <button
              type="button"
              onClick={() => onChangeTarget('inventory')}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                target === 'inventory'
                  ? 'bg-[#10312B] text-white'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Inventario
            </button>
            <button
              type="button"
              onClick={() => onChangeTarget('full')}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                target === 'full'
                  ? 'bg-[#10312B] text-white'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Reporte completo
            </button>
          </div>
        </div>

        {/* Modal Body: Image Preview */}
        <div className="flex-1 p-5 overflow-y-auto bg-gray-100 flex items-center justify-center min-h-[300px]">
          {loading ? (
            <div className="text-center py-12 flex flex-col items-center gap-3">
              <Loader2 className="w-8 h-8 animate-spin text-[#691C32]" />
              <div className="text-xs font-semibold text-gray-700">
                Generando captura de alta resolución ({getTargetTitle()})...
              </div>
              <p className="text-[11px] text-gray-500">Optimizando colores y tipografía institucional</p>
            </div>
          ) : errorMessage ? (
            <div className="text-center py-8 text-rose-700 text-xs">
              <p className="font-bold">Error al generar captura:</p>
              <p className="mt-1">{errorMessage}</p>
            </div>
          ) : imageDataUrl ? (
            <div className="border border-gray-300 rounded shadow-sm overflow-hidden bg-white max-w-full">
              <img
                src={imageDataUrl}
                alt="Vista previa de reporte"
                className="w-full h-auto max-h-[460px] object-contain"
              />
            </div>
          ) : null}
        </div>

        {/* Modal Footer Controls */}
        <div className="px-5 py-3.5 border-t border-gray-200 bg-white flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
          <div className="text-gray-500 text-[11px]">
            Formato: <strong>PNG 2x</strong> · Fondo: <strong>Blanco institucional</strong>
          </div>

          <div className="flex items-center gap-2 self-end sm:self-auto">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-2 rounded border border-gray-300 text-gray-700 hover:bg-gray-50 font-medium transition-colors"
            >
              Cancelar
            </button>

            {/* Download Button */}
            <button
              type="button"
              disabled={loading || !imageDataUrl}
              onClick={handleDownload}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded bg-[#10312B] hover:bg-[#0c2420] text-white font-semibold transition-colors disabled:opacity-50 shadow-xs"
            >
              <Download className="w-4 h-4" />
              <span>Descargar PNG</span>
            </button>

            {/* Share Button */}
            <button
              type="button"
              disabled={loading || !imageBlob}
              onClick={handleShare}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded bg-[#691C32] hover:bg-[#521426] text-white font-semibold transition-colors disabled:opacity-50 shadow-xs"
            >
              {shareSuccess ? (
                <>
                  <Check className="w-4 h-4 text-emerald-300" />
                  <span>¡Compartido!</span>
                </>
              ) : (
                <>
                  <MessageCircle className="w-4 h-4" />
                  <span>Enviar PNG</span>
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
