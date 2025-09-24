import { useState, useRef } from "react"

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";


interface Wine {
    name: string,
    nation: string,
    region: string,
    category: string,
    volume: number,
    normalPrice: number,
    nowPrice: number
}


const WineComponent = () => {

    const fileInputRef = useRef<HTMLInputElement | null>(null);
    const [datas, setDatas] = useState<Array<Wine>>([])

    
    const getWineBlockPosition = (index: number) => {
        const colStarts = [2, 9, 16, 23]; // B, I, P, W
        const blockWidth = 6;   // G-B+1
        const blockHeight = 8;  // 행 높이

        const rowSpacing = 9;   // 8행 + 1행 간격
        const group = Math.floor(index / 4); // 몇 번째 줄인가
        const posInRow = index % 4;          // 줄 안에서 몇 번째 열인가

        const startRow = 2 + group * rowSpacing;
        const startCol = colStarts[posInRow];

        const endRow = startRow + blockHeight - 1; // 8행짜리
        const endCol = startCol + blockWidth - 1;  // 6열짜리

        return { startRow, endRow, startCol, endCol };
    }

    const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (event) => {
            const buffer = event.target?.result;
            if (!buffer) return;

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer as ArrayBuffer);

            const sheet = workbook.worksheets[0]; // 첫 번째 시트 사용


            const wines: Wine[] = [];

            // 예: 2행부터 데이터 있다고 가정
            sheet.eachRow((row, rowNumber) => {


                if (rowNumber === 1) return; // 헤더 건너뛰기

                const [
                    header,
                    name,
                    nation,
                    region,
                    category,
                    volume,
                    normalPrice,
                    nowPrice,
                ] = row.values as Array<any>;


                if (!name) return;

                wines.push({
                    name,
                    nation,
                    region,
                    category,
                    volume: Number(volume),
                    normalPrice: Number(normalPrice),
                    nowPrice: Number(nowPrice),
                });
            });

            console.log("setDatas")
            console.log(wines)

            setDatas(wines);
        };

        reader.readAsArrayBuffer(file);
    };

    const handleExport = async () => {

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("와인 라벨");

        sheet.getRow(1).height = 5;
        sheet.getColumn(1).width = 1;


        let flags: Array<{startRow: number, startCol: number, nation: string}> = []

        datas.forEach(async (wine, idx) => {
            const { startRow, endRow, startCol, endCol } = getWineBlockPosition(idx);

            flags.push({startRow: startRow, startCol: startCol, nation: wine.nation})

            // 외곽 테두리
            for (let r = startRow; r <= endRow; r++) {
                for (let c = startCol; c <= endCol; c++) {
                    const cell = sheet.getCell(r, c);

                    if (r === startRow) cell.border = { ...cell.border, top: { style: "thin" } };
                    if (r === endRow) cell.border = { ...cell.border, bottom: { style: "thin" } };
                    if (c === startCol) cell.border = { ...cell.border, left: { style: "thin" } };
                    if (c === endCol) cell.border = { ...cell.border, right: { style: "thin" } };
                }
            }

            // C3~E4 대신 상대 위치
            sheet.mergeCells(startRow + 1, startCol + 1, startRow + 2, startCol + 3);
            const nameCell = sheet.getCell(startRow + 1, startCol + 1);
            nameCell.value = wine.name;
            nameCell.alignment = { horizontal: "left", vertical: "middle" };
            nameCell.font = { bold: true, size: 19 };

            // 국가 > 지역
            sheet.mergeCells(startRow + 3, startCol + 1, startRow + 3, startCol + 3);
            sheet.getCell(startRow + 3, startCol + 1).value = `${wine.nation} > ${wine.region}`;

            // category
            sheet.mergeCells(startRow + 4, startCol + 1, startRow + 4, startCol + 3);
            sheet.getCell(startRow + 4, startCol + 1).value = wine.category;

            // volume
            sheet.getCell(startRow + 4, startCol + 4).value = `${wine.volume}ml`;

            // 정상가

            sheet.mergeCells(startRow + 5, startCol + 1, startRow + 5, startCol + 2);
            const normalPriceTitleCell = sheet.getCell(startRow + 5, startCol + 2);
            normalPriceTitleCell.value = "정상가";
            normalPriceTitleCell.alignment = { vertical: "middle", horizontal: "center" };


            sheet.mergeCells(startRow + 6, startCol + 1, startRow + 6, startCol + 2);
            const normalPriceCell = sheet.getCell(startRow + 6, startCol + 2);
            normalPriceCell.value = `${wine.normalPrice}원`;
            normalPriceCell.font = { strike: true, };
            normalPriceCell.alignment = { vertical: "middle", horizontal: "center" };

            // nowPrice
            sheet.mergeCells(startRow + 5, startCol + 3, startRow + 6, startCol + 4);
            const nowPriceCell = sheet.getCell(startRow + 5, startCol + 3);
            nowPriceCell.value = `${wine.nowPrice}원`;
            nowPriceCell.alignment = { vertical: "middle", horizontal: "center" };
            nowPriceCell.font = { bold: true, size: 18, color: { argb: "FF0000" } };

        });

        for (let i = 0; i<flags.length; i++){

            const flag = flags[i]

            let fileName = ""
            
            if(flag.nation === '이탈리아'){
                fileName = "italy"
            }else if(flag.nation === '프랑스'){
                fileName = "france"
            }

            const response = await fetch(process.env.PUBLIC_URL + `/imgs/${fileName}.jpeg`);
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();

            console.log(arrayBuffer)

            const imageId = workbook.addImage({
                buffer: arrayBuffer,
                extension: "jpeg",
            })

            sheet.addImage(imageId, {
                tl: { col: flags[i].startCol + 3, row: flags[i].startRow },
                ext: { width: 74, height: 96 }, // px 단위 크기 지정
                editAs: "absolute"
            });
        }
           

        // 엑셀 파일로 다운로드
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "wine.xlsx");
    }


    const downloadTemplate = () => {
        const link = document.createElement("a");
        link.href = process.env.PUBLIC_URL + "/xlsx/wine_label_template.xlsx";
        link.download = "wine_template.xlsx";
        link.click();
    };

    return <>
        Wine 네익택 생성기 <br/>
        <button onClick={()=>{
            fileInputRef.current?.click();
        }}>엑셀 업로드</button>
            <input type="file"
                   accept=".xlsx,.xls"
                   ref={fileInputRef}
                   style={{ display: "none" }}
                   onChange={handleImport}/>
        <br/>

        <button onClick={downloadTemplate}>템플릿 다운로드</button><br/>

        now wine datas : {datas.length}
        <br/>
        <button onClick={handleExport}>엑셀로 내보내기</button>
    </>
}

export default WineComponent