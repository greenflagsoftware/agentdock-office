import { useEffect, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
// Vite resolves this to a URL for the worker asset at build time.
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

interface PdfViewerProps {
  url: string;
}

export default function PdfViewer({ url }: PdfViewerProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const docRef = useRef<pdfjsLib.PDFDocumentProxy | null>(null);
  const [pageNum, setPageNum] = useState(1);
  const [numPages, setNumPages] = useState(0);
  const [error, setError] = useState<string | null>(null);

  // Load the document whenever the source URL changes.
  useEffect(() => {
    let cancelled = false;
    setError(null);
    setPageNum(1);
    setNumPages(0);
    docRef.current = null;

    pdfjsLib
      // standardFontDataUrl is required whenever a PDF uses a non-embedded
      // base-14 font (Helvetica, Times, etc.) — without it, page.render()
      // hangs forever waiting on glyph data that never arrives.
      .getDocument({ url, standardFontDataUrl: "/standard_fonts/" })
      .promise.then((doc) => {
        if (cancelled) return;
        docRef.current = doc;
        setNumPages(doc.numPages);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : "Failed to load PDF.");
      });

    return () => {
      cancelled = true;
      docRef.current?.destroy();
    };
  }, [url]);

  // Render the current page whenever it changes (or the document finishes loading).
  useEffect(() => {
    let cancelled = false;
    const doc = docRef.current;
    const canvas = canvasRef.current;
    if (!doc || !canvas || numPages === 0) return;

    doc.getPage(pageNum).then((page) => {
      if (cancelled) return;
      const viewport = page.getViewport({ scale: 1.5 });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const context = canvas.getContext("2d");
      if (!context) return;
      page.render({ canvasContext: context, viewport }).promise.catch((err: unknown) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : "Failed to render page.");
      });
    }).catch((err: unknown) => {
      if (cancelled) return;
      setError(err instanceof Error ? err.message : "Failed to render page.");
    });

    return () => {
      cancelled = true;
    };
  }, [pageNum, numPages]);

  if (error) {
    return <div className="pdf-viewer-error">Could not display PDF: {error}</div>;
  }

  return (
    <div className="pdf-viewer">
      {numPages > 1 && (
        <div className="pdf-viewer-controls">
          <button
            className="btn btn-sm"
            onClick={() => setPageNum((p) => Math.max(1, p - 1))}
            disabled={pageNum <= 1}
          >
            Previous
          </button>
          <span className="pdf-viewer-page-indicator">
            Page {pageNum} of {numPages}
          </span>
          <button
            className="btn btn-sm"
            onClick={() => setPageNum((p) => Math.min(numPages, p + 1))}
            disabled={pageNum >= numPages}
          >
            Next
          </button>
        </div>
      )}
      <div className="pdf-viewer-canvas-wrap">
        <canvas ref={canvasRef} />
      </div>
    </div>
  );
}
