import { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Loader2, FileSpreadsheet, Trash2, Camera, Clipboard } from 'lucide-react';

type Fila = {
  'Fecha Ingreso': string;
  Local: string;
  'Código': string;
  'Descripción': string;
  'Caja': string;
  'Palmeta': string;
  'Unidad': string;
  'Nota de Venta': string;
  'Guía de Paso': string;
  'Fecha Guía de Paso': string;
};

const COLUMNAS: (keyof Fila)[] = [
  'Fecha Ingreso',
  'Local',
  'Código',
  'Descripción',
  'Caja',
  'Palmeta',
  'Unidad',
  'Nota de Venta',
  'Guía de Paso',
  'Fecha Guía de Paso',
];

function parsearFormulario(textoOCR: string): Fila {
  const fila: Fila = {
    'Fecha Ingreso': new Date().toLocaleDateString('es-CL'),
    Local: '',
    'Código': '',
    'Descripción': '',
    'Caja': '',
    'Palmeta': '',
    'Unidad': '',
    'Nota de Venta': '',
    'Guía de Paso': '',
    'Fecha Guía de Paso': '',
  };

  const lineas = textoOCR
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const norm = (s: string) =>
    s
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]/g, '')
      .toLowerCase();

  const idx = (etiquetaNorm: string) => lineas.findIndex((l) => norm(l) === etiquetaNorm);

  const idxLocal = idx('local');
  const idxCodigo = idx('codigo');
  const idxDescripcion = idx('descripcion');
  const idxCaja = idx('caja');
  const idxPalmeta = idx('palmeta');
  const idxUnidad = idx('unidad');
  const idxNota = idx('notadeventa');
  const idxGuia = idx('guiadepaso');
  const idxFecha = idx('fechaguiadepaso');

  if (idxLocal !== -1) fila.Local = lineas[idxLocal + 1] ?? '';
  if (idxCodigo !== -1) fila['Código'] = lineas[idxCodigo + 1] ?? '';

  if (idxDescripcion !== -1) {
    const antes = lineas[idxDescripcion - 1] ?? '';
    const despues = lineas[idxDescripcion + 1] ?? '';
    const partes: string[] = [];
    if (antes && antes !== fila['Código'] && norm(antes) !== 'codigo') partes.push(antes);
    const etiquetasDesc = ['cantidad', 'caja', 'palmeta', 'unidad', 'notadeventa', 'guiadepaso', 'fechaguiadepaso'];
    if (despues && !etiquetasDesc.includes(norm(despues))) partes.push(despues);
    fila['Descripción'] = partes.join(' ').trim();
  }

  const tomarNumeroDespues = (i: number): string => {
    if (i === -1) return '';
    const val = lineas[i + 1] ?? '';
    return /^\d+$/.test(val) ? val : '';
  };

  fila['Caja'] = tomarNumeroDespues(idxCaja);
  fila['Palmeta'] = tomarNumeroDespues(idxPalmeta);
  fila['Unidad'] = tomarNumeroDespues(idxUnidad);

  if (idxNota !== -1) fila['Nota de Venta'] = (lineas[idxNota + 1] ?? '').replace(/[^0-9]/g, '');
  if (idxGuia !== -1) fila['Guía de Paso'] = (lineas[idxGuia + 1] ?? '').replace(/[^0-9]/g, '');
  if (idxFecha !== -1) {
    const val = lineas[idxFecha + 1] ?? '';
    const m = val.match(/\d{2}[-/]\d{2}[-/]\d{2,4}/);
    fila['Fecha Guía de Paso'] = m ? m[0] : val;
  }

  return fila;
}

export default function App() {
  const [imagen, setImagen] = useState<string | null>(null);
  const [filas, setFilas] = useState<Fila[]>([]);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [textoOCR, setTextoOCR] = useState<string>('');
  const [pegadoReciente, setPegadoReciente] = useState(false);

  const manejarArchivo = (file: File) => {
    setError(null);
    setTextoOCR('');
    const reader = new FileReader();
    reader.onload = (e) => setImagen(e.target?.result as string);
    reader.readAsDataURL(file);
  };

  // Detectar Ctrl+V en cualquier parte de la página
  const manejarPegado = useCallback((e: ClipboardEvent) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.startsWith('image/')) {
        const file = items[i].getAsFile();
        if (file) {
          manejarArchivo(file);
          setPegadoReciente(true);
          setTimeout(() => setPegadoReciente(false), 2000);
          e.preventDefault();
        }
        return;
      }
    }
  }, []);

  useEffect(() => {
    window.addEventListener('paste', manejarPegado as any);
    return () => window.removeEventListener('paste', manejarPegado as any);
  }, [manejarPegado]);

  const extraerDatos = async () => {
    if (!imagen) return;
    setCargando(true);
    setError(null);
    setTextoOCR('');
    try {
      const ffocr = await import('ffocr');

      const resp = await fetch(imagen);
      const blob = await resp.blob();
      const file = new File([blob], 'formulario.jpg', { type: blob.type });

      const { createDefaultPPOcrV5 } = ffocr as any;
      const ocr = createDefaultPPOcrV5();
      const resultado = await ocr.ocr(file);
      const texto = resultado?.text ?? resultado?.data?.text ?? String(resultado);
      setTextoOCR(texto);

      const nuevaFila = parsearFormulario(texto);
      setFilas((prev) => [...prev, nuevaFila]);
      setImagen(null);
    } catch (e: any) {
      console.error(e);
      setError(
        'No se pudo procesar la imagen. Detalle: ' +
          (e?.message || 'error desconocido') +
          '. Revisa la consola (F12) para más detalles.'
      );
    } finally {
      setCargando(false);
    }
  };

  const descargarExcel = () => {
    const ordenado = filas.map((f) => {
      const obj: Record<string, any> = {};
      COLUMNAS.forEach((c) => (obj[c] = f[c] ?? ''));
      return obj;
    });
    const ws = XLSX.utils.json_to_sheet(ordenado, { header: COLUMNAS as string[] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reposiciones');
    XLSX.writeFile(wb, `reposiciones_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const limpiarTodo = () => {
    setFilas([]);
    setImagen(null);
    setError(null);
    setTextoOCR('');
  };

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      <div className="max-w-5xl mx-auto">
        <h1 className="text-3xl font-bold text-slate-800 mb-2">
          Lector de Reposiciones Limatco
        </h1>
        <p className="text-slate-600 mb-8">
          Pega un recorte con <kbd className="px-2 py-0.5 bg-slate-200 rounded text-xs font-mono">Ctrl + V</kbd>,
          o sube una foto. El OCR local leerá los datos.
        </p>

        <div
          onDrop={(e) => {
            e.preventDefault();
            const f = e.dataTransfer.files[0];
            if (f) manejarArchivo(f);
          }}
          onDragOver={(e) => e.preventDefault()}
          className="border-2 border-dashed border-slate-300 rounded-xl p-10 text-center bg-white hover:border-blue-400 transition cursor-pointer"
        >
          <input
            type="file"
            accept="image/*"
            id="file-input"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) manejarArchivo(f);
              e.target.value = '';
            }}
          />
          <label htmlFor="file-input" className="cursor-pointer block">
            <Upload className="w-12 h-12 mx-auto text-slate-400 mb-3" />
            <p className="text-slate-700 font-medium">
              Arrastra la imagen aquí, haz clic para seleccionar, o pega con Ctrl + V
            </p>
            <p className="text-slate-400 text-sm mt-1 flex items-center justify-center gap-1">
              <Clipboard className="w-3.5 h-3.5" />
              Consejo: usa Windows + Shift + S para recortar y luego Ctrl + V aquí
            </p>
          </label>
        </div>

        {pegadoReciente && !imagen && (
          <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg text-green-700 text-sm">
            ✓ Imagen pegada correctamente
          </div>
        )}

        {imagen && (
          <div className="mt-6 bg-white rounded-xl p-4 shadow-sm">
            <img src={imagen} alt="Formulario" className="max-h-96 mx-auto rounded" />
            <button
              onClick={extraerDatos}
              disabled={cargando}
              className="mt-4 w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 flex items-center justify-center gap-2"
            >
              {cargando ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" /> Procesando imagen (la primera vez puede tardar)...
                </>
              ) : (
                <>
                  <Camera className="w-5 h-5" /> Extraer datos
                </>
              )}
            </button>
          </div>
        )}

        {error && (
          <div className="mt-4 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
            {error}
          </div>
        )}

        {textoOCR && (
          <details className="mt-4 bg-white rounded-xl p-4 shadow-sm">
            <summary className="cursor-pointer text-slate-600 font-medium">
              Ver texto detectado por OCR (para depuración)
            </summary>
            <pre className="mt-3 p-3 bg-slate-100 rounded text-xs overflow-auto whitespace-pre-wrap">
              {textoOCR}
            </pre>
          </details>
        )}

        {filas.length > 0 && (
          <div className="mt-6 bg-white rounded-xl p-4 shadow-sm">
            <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
              <h2 className="text-lg font-semibold text-slate-800">
                Datos extraídos ({filas.length} {filas.length === 1 ? 'formulario' : 'formularios'})
              </h2>
              <div className="flex gap-2">
                <button
                  onClick={limpiarTodo}
                  className="bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-medium hover:bg-slate-300 flex items-center gap-2"
                >
                  <Trash2 className="w-4 h-4" /> Limpiar
                </button>
                <button
                  onClick={descargarExcel}
                  className="bg-green-600 text-white px-4 py-2 rounded-lg font-medium hover:bg-green-700 flex items-center gap-2"
                >
                  <FileSpreadsheet className="w-4 h-4" /> Descargar Excel
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-slate-100">
                  <tr>
                    {COLUMNAS.map((k) => (
                      <th key={k} className="px-3 py-2 text-left font-medium text-slate-700 whitespace-nowrap">
                        {k}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filas.map((fila, i) => (
                    <tr key={i} className="border-t border-slate-100">
                      {COLUMNAS.map((c) => (
                        <td key={c} className="px-3 py-2 text-slate-600 whitespace-nowrap">
                          {fila[c] || '-'}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}