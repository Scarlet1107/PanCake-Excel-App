'use client';

import { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import ExcelJS from 'exceljs';
import { NotesText } from './constants';

export default function HomePage() {
  const [message, setMessage] = useState<string | null>(null);
  const writeNotesToCell = (worksheet: ExcelJS.Worksheet) => {
    NotesText.forEach(([col, row, text]) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.value = text;
    });
  };

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    setMessage(null);
    const file = acceptedFiles[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        const worksheet = workbook.getWorksheet('Sheet1');
        if (!worksheet) {
          setMessage('Sheet1 が存在しません。');
          return;
        }

        let row = 5;
        let processedCount = 0;

        while (true) {
          const cell = worksheet.getCell(`H${row}`);
          const value = cell.value;

          if (typeof value !== 'number' || Number.isNaN(value)) {
            break;
          }

          // 0.9倍して四捨五入、同じセルに上書き
          cell.value = Math.round(value * 0.9);

          row++;
          processedCount++;
        }

        // 注意事項を書き込む
        writeNotesToCell(worksheet);

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        // a.download = 'rounded_modified.xlsx';
        a.download = "(変換済み)"+file.name
        a.click();
        URL.revokeObjectURL(url);

        setMessage(`仕入れ単価の数値を変換して注意事項を入力しました（${processedCount} 件処理）`);
      } catch (error) {
        console.error(error);
        setMessage('処理中にエラーが発生しました。');
      }
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
    },
    multiple: false,
  });

  return (
    <div className="flex items-center justify-center min-h-screen bg-gray-100 p-4">
      <div className="text-center space-y-6">
        <div
          {...getRootProps()}
          className={`border-4 border-dashed rounded-xl p-10 bg-white shadow-md cursor-pointer transition-colors duration-300
            ${isDragActive ? 'border-blue-400 bg-blue-50' : 'border-gray-300'}
          `}
        >
          <input {...getInputProps()} />
          <p className="text-gray-600 text-lg">
            Excelファイル（.xlsx）をここにドラッグ＆ドロップ
          </p>
        </div>
        {message && (
          <p className="text-red-500 font-semibold">{message}</p>
        )}
      </div>
    </div>
  );
}
