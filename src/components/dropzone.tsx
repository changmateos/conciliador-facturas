"use client";

import { useRef, useState } from "react";

interface DropzoneProps {
  title: string;
  subtitle: string;
  multiple: boolean;
  onFiles: (files: File[]) => void;
}

export default function Dropzone({ title, subtitle, multiple, onFiles }: DropzoneProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  function accept(list: FileList | null) {
    if (!list) return;
    const valid = Array.from(list).filter((f) => /\.(xlsx|xls|csv)$/i.test(f.name));
    if (valid.length > 0) onFiles(valid);
  }

  return (
    <div
      role="button"
      tabIndex={0}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(e) => { if (e.key === "Enter") inputRef.current?.click(); }}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => { e.preventDefault(); setDragging(false); accept(e.dataTransfer.files); }}
      className={
        "w-full cursor-pointer rounded-xl border-2 border-dashed p-6 text-left transition-colors " +
        (dragging
          ? "border-blue-600 bg-blue-50"
          : "border-gray-300 bg-white hover:border-blue-500 hover:bg-blue-50/40")
      }
    >
      <div className="flex items-center gap-4">
        <span className="w-10 h-10 rounded-lg bg-blue-700 text-white flex items-center justify-center shrink-0">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0l-4 4m4-4l4 4M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3" />
          </svg>
        </span>
        <span className="min-w-0">
          <span className="block text-sm font-bold text-gray-900">{title}</span>
          <span className="block text-[12px] text-gray-500 mt-0.5">{subtitle}</span>
        </span>
      </div>
      <input
        ref={inputRef}
        type="file"
        multiple={multiple}
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onClick={(e) => e.stopPropagation()}
        onChange={(e) => { accept(e.target.files); e.target.value = ""; }}
      />
    </div>
  );
}