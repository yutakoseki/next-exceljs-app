"use client";
import React from "react";
import ExcelJS from "exceljs";

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
    // handleFileChange関数の修正
    const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files && event.target.files[0];

        if (file) {
            try {
                const buffer = await fileToBuffer(file);
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);

                // ここでExcelファイルを処理するためのコードを記述
                // 新しいシートを作成し、データを追加
                const newSheet = workbook.addWorksheet("新規");
                newSheet.addRow([1, "New Data", 42]);

                // Excelファイルを加工した後の処理を記述
                // 例: 加工したファイルをダウンロード
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
            }
        }
    };

    return (
        <div>
            <input type="file" accept=".xlsx" onChange={(event) => handleFileChange(event)} />
        </div>
    );
};

export default UploadExcelComponent;
