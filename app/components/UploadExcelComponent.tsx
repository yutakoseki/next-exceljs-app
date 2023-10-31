"use client";
import React from "react";
import ExcelJS, { Workbook, Worksheet } from "exceljs";

const UploadExcelComponent: React.FC = () => {
    // FileからBufferへの変換関数
    async function fileToBuffer(file: File): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                if (event.target) {
                    const arrayBuffer = event.target.result as ArrayBuffer;
                    const buffer = Buffer.from(arrayBuffer);
                    resolve(buffer);
                } else {
                    reject(new Error("Error reading file"));
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // 集計のための関数
    const aggregateData = (workbook: Workbook) => {
        // シートの追加
        const lunchSheet = workbook.addWorksheet("詳細　昼");
        const snackSheet = workbook.addWorksheet("詳細　おやつ");
        const sumLunchSheet = workbook.addWorksheet("集計　昼");
        const sumSnackSheet = workbook.addWorksheet("集計　おやつ");

        // 表題行に項目を追加
        lunchSheet.addRow(["材料名", "3-5歳児分量"]);
        snackSheet.addRow(["材料名", "3-5歳児分量"]);
        sumLunchSheet.addRow(["材料名", "3-5歳児分量"]);
        sumSnackSheet.addRow(["材料名", "3-5歳児分量"]);

        const dataLunch: { colC: any; colE: any }[] = [];
        const dataSnackCE: { colC: any; colE: any }[] = [];
        const dataSnackPR: { colP: any; colR: any }[] = [];

        workbook.eachSheet((worksheet: Worksheet) => {
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                // 昼の行のみ抽出
                if (row.getCell(1).value === "昼") {
                    const colCValue = row.getCell(16).value; // お手持ちの調味料
                    const colEValue = row.getCell(18).value; // 3-5歳児分量

                    if (colCValue !== null && colEValue !== null) {
                        dataLunch.push({ colC: colCValue, colE: colEValue });
                    }
                }
                // おやつの行のみ抽出
                if (row.getCell(1).value === "おやつ") {
                    const colCValue = row.getCell(3).value;
                    const colEValue = row.getCell(5).value;
                    const colPValue = row.getCell(16).value;
                    const colRValue = row.getCell(18).value;

                    if (colCValue !== null && colEValue !== null) {
                        dataSnackCE.push({ colC: colCValue, colE: colEValue });
                    }

                    if (colPValue !== null && colRValue !== null) {
                        dataSnackPR.push({ colP: colPValue, colR: colRValue });
                    }
                }
            });
        });

        if (dataLunch.length > 0) {
            lunchSheet.addRows(dataLunch.map((item) => [item.colC, item.colE]));
        }
        if (dataSnackCE.length > 0) {
            snackSheet.addRows(dataSnackCE.map((item) => [item.colC, item.colE]));
        }
        if (dataSnackPR.length > 0) {
            snackSheet.addRows(dataSnackPR.map((item) => [item.colP, item.colR]));
        }

        // 同じ材料名を持つ行を探して合計する
        const aggregateDataByKey = (data: { [key: string]: number }, dataArray: { colC: string; colE: number }[]) => {
            dataArray.forEach((item) => {
                if (!data[item.colC]) {
                    data[item.colC] = item.colE;
                } else {
                    data[item.colC] += item.colE;
                }
            });
            return data;
        };

        if (dataLunch.length > 0) {
            const aggregatedLunchData: { [key: string]: number } = {};
            const aggregatedDataLunch = aggregateDataByKey(aggregatedLunchData, dataLunch);
            Object.keys(aggregatedDataLunch).forEach((key) => {
                sumLunchSheet.addRow([key, aggregatedDataLunch[key]]);
            });
        }

        if (dataSnackCE.length > 0) {
            const aggregatedSnackCEData: { [key: string]: number } = {};
            const aggregatedDataSnackCE = aggregateDataByKey(aggregatedSnackCEData, dataSnackCE);
            Object.keys(aggregatedDataSnackCE).forEach((key) => {
                sumSnackSheet.addRow([key, aggregatedDataSnackCE[key]]);
            });
        }

        if (dataSnackPR.length > 0) {
            const aggregatedSnackPRData: { [key: string]: number } = {};
            dataSnackPR.forEach((item) => {
                const key = item.colP;
                const value = item.colR;
                if (!aggregatedSnackPRData[key]) {
                    aggregatedSnackPRData[key] = value;
                } else {
                    aggregatedSnackPRData[key] += value;
                }
            });
            Object.keys(aggregatedSnackPRData).forEach((key) => {
                snackSheet.addRow([key, aggregatedSnackPRData[key]]);
            });
        }
    };

    // ファイル変更時の処理
    const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files && event.target.files[0];

        if (file) {
            try {
                const buffer = await fileToBuffer(file);
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);

                aggregateData(workbook);

                // Excelファイルを加工した後の処理
                const bufferToWrite = await workbook.xlsx.writeBuffer();
                const blob = new Blob([bufferToWrite], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                const url = URL.createObjectURL(blob);

                const link = document.createElement("a");
                link.href = url;
                link.setAttribute("download", "processed_file.xlsx");
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            } catch (error) {
                console.error("Error processing Excel file:", error);
                // エラーメッセージをユーザーに表示（必要に応じて）
            }
        }
    };

    return (
        <div className="flex items-center justify-center h-screen">
          <label htmlFor="file-upload" className="flex flex-col items-center px-4 py-6 bg-gray-100 rounded-md shadow-md tracking-wide cursor-pointer">
            <svg
              className="w-8 h-8 mb-2"
              fill="currentColor"
              xmlns="http://www.w3.org/2000/svg"
              viewBox="0 0 20 20"
            >
              <path
                fillRule="evenodd"
                d="M10 3a1 1 0 0 1 1 1v4.586l2.293-2.293a1 1 0 1 1 1.414 1.414l-4 4a1 1 0 0 1-1.414 0l-4-4a1 1 0 1 1 1.414-1.414L9 8.586V4a1 1 0 0 1 1-1z"
              />
              <path
                fillRule="evenodd"
                d="M2 13a1 1 0 0 1 1-1h14a1 1 0 0 1 0 2H3a1 1 0 0 1-1-1z"
              />
            </svg>
            <span className="text-base leading-normal">ファイルを選択</span>
            <input
              id="file-upload"
              type="file"
              accept=".xlsx"
              onChange={(event) => handleFileChange(event)}
              className="hidden"
            />
          </label>
        </div>
      );
};

export default UploadExcelComponent;
