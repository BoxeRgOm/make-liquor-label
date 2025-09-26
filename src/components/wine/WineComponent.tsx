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
        const blockWidth = 6;   // G-B+10
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

    const [widthColumns, setWidthColumns] = useState({
        space : 2.2,
        colums0 : 0.41,
        colums1 : 14,
        colums2 : 7,
        colums3 : 7.5,
        colums4 : 5,
        colums5 : 0.41,
    })

    const [heightRows, setHeightRows] = useState({
        row0: 5,
        row1: 50,
        row2: 15,
        row3: 15,
        row4: 5,
        row5: 17.4,
        row6: 23.4,
        row7: 5,
    })

    const [sizeFlag, setSizeFlag] = useState({
         width: 42, height: 80 
        })

    const handleExport = async () => {

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("와인 라벨");

        sheet.getRow(1).height = 2.2;

        sheet.getColumn(1).width = widthColumns.space;

        sheet.getColumn(2).width = widthColumns.colums0;
        sheet.getColumn(3).width = widthColumns.colums1;
        sheet.getColumn(4).width = widthColumns.colums2;
        sheet.getColumn(5).width = widthColumns.colums3;
        sheet.getColumn(6).width = widthColumns.colums4;
        sheet.getColumn(7).width = widthColumns.colums5;

        sheet.getColumn(8).width = widthColumns.space;

        sheet.getColumn(9).width = widthColumns.colums0;
        sheet.getColumn(10).width = widthColumns.colums1;
        sheet.getColumn(11).width = widthColumns.colums2;
        sheet.getColumn(12).width = widthColumns.colums3;
        sheet.getColumn(13).width = widthColumns.colums4;
        sheet.getColumn(14).width = widthColumns.colums5;

        sheet.getColumn(15).width = widthColumns.space;

        sheet.getColumn(16).width = widthColumns.colums0;
        sheet.getColumn(17).width = widthColumns.colums1;
        sheet.getColumn(18).width = widthColumns.colums2;
        sheet.getColumn(19).width = widthColumns.colums3;
        sheet.getColumn(20).width = widthColumns.colums4;
        sheet.getColumn(21).width = widthColumns.colums5;

        sheet.getColumn(22).width = widthColumns.space;

        sheet.getColumn(23).width = widthColumns.colums0;
        sheet.getColumn(24).width = widthColumns.colums1;
        sheet.getColumn(25).width = widthColumns.colums2;
        sheet.getColumn(26).width = widthColumns.colums3;
        sheet.getColumn(27).width = widthColumns.colums4;
        sheet.getColumn(28).width = widthColumns.colums5;



        let flags: Array<{startRow: number, startCol: number, nation: string}> = []

        datas.forEach(async (wine, idx) => {
            const { startRow, endRow, startCol, endCol } = getWineBlockPosition(idx);

            flags.push({startRow: startRow, startCol: startCol, nation: wine.nation})

            // 외곽 테두리
            for (let r = startRow; r <= endRow; r++) {
                for (let c = startCol; c <= endCol; c++) {
                    const cell = sheet.getCell(r, c);

                    if (r === startRow) cell.border = { ...cell.border, top: { style: "medium" } };
                    if (r === endRow) cell.border = { ...cell.border, bottom: { style: "medium" } };
                    if (c === startCol) cell.border = { ...cell.border, left: { style: "medium" } };
                    if (c === endCol) cell.border = { ...cell.border, right: { style: "medium" } };
                }
            }

            if (idx % 4 === 0){

                sheet.getRow(startRow).height = heightRows.row0;
                sheet.getRow(startRow + 1).height = heightRows.row1;
                sheet.getRow(startRow + 2).height = heightRows.row2;
                sheet.getRow(startRow + 3).height = heightRows.row3;
                sheet.getRow(startRow + 4).height = heightRows.row4;
                sheet.getRow(startRow + 5).height = heightRows.row5;
                sheet.getRow(startRow + 6).height = heightRows.row6;
                sheet.getRow(startRow + 7).height = heightRows.row7;

            }

            // C3~E4 대신 상대 위치
            sheet.mergeCells(startRow + 1, startCol + 1, startRow + 1, startCol + 3);

            const nameCell = sheet.getCell(startRow + 1, startCol + 1);
            nameCell.value = wine.name;
            nameCell.alignment = { wrapText: true, horizontal: "left", vertical: "middle" };
            nameCell.font = {bold: true, size: 19 };

            // 국가 > 지역
            sheet.mergeCells(startRow + 2, startCol + 1, startRow + 2, startCol + 2);
            const regionCell = sheet.getCell(startRow + 2, startCol + 1);
            if (wine.region){
                regionCell.value = `${wine.nation} > ${wine.region}`;
            }else{
                regionCell.value = `${wine.nation}`;
            }
            regionCell.font = { size: 8 };


            // category
            sheet.mergeCells(startRow + 3, startCol + 1, startRow + 3, startCol + 3);
            const categoryCell = sheet.getCell(startRow + 3, startCol + 1);
            categoryCell.value = wine.category;
            categoryCell.font = { size: 8 };

            // volume
            const volumeCell = sheet.getCell(startRow + 3, startCol + 4);
            volumeCell.value = `${wine.volume}ml`;
            volumeCell.font = { size: 8 };

            // 정상가
            const normalPriceTitleCell = sheet.getCell(startRow + 5, startCol + 1);
            normalPriceTitleCell.value = "정상가";
            normalPriceTitleCell.font = { size: 8 };
            normalPriceTitleCell.alignment = { vertical: "middle", horizontal: "center" };

            const normalPriceCell = sheet.getCell(startRow + 6, startCol + 1);

            normalPriceCell.value = `${wine.normalPrice.toLocaleString()}원`;
            normalPriceCell.font = { strike: true, size: 12, bold: true};
            normalPriceCell.alignment = { vertical: "middle", horizontal: "center" };

            // nowPrice
            sheet.mergeCells(startRow + 5, startCol + 2, startRow + 6, startCol + 4);
            const nowPriceCell = sheet.getCell(startRow + 5, startCol + 2);

            nowPriceCell.value = `${wine.nowPrice.toLocaleString()}원`;
            nowPriceCell.alignment = { vertical: "middle", horizontal: "center" };
            nowPriceCell.font = { bold: true, size: 18, color: { argb: "FF0000" } };

        });

        for (let i = 0; i<flags.length; i++){

            const flag = flags[i]

            let fileName = ""
            
            if(flag.nation === '프랑스'){
                fileName = "fr"
            }else if(flag.nation === '이탈리아'){
                fileName = "it"
            }else if(flag.nation === '독일'){
                fileName = "dc"
            }else if(flag.nation === '칠레'){
                fileName = "ch"
            }else if(flag.nation === '호주'){
                fileName = "aus"
            }else if(flag.nation === '스페인'){
                fileName = "dc"
            }else if(flag.nation === '포르투갈'){
                fileName = "pg"
            }else if(flag.nation === '미국'){
                fileName = "us"
            }else if(flag.nation === '뉴질랜드'){
                fileName = "nz"
            }else if(flag.nation === '남아프리카공화국'){
                fileName = "sa"
            }else if(flag.nation === '아르헨티나'){
                fileName = "ag"
            }
            const response = await fetch(process.env.PUBLIC_URL + `/imgs/${fileName}.png`);
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();

            console.log(arrayBuffer)

            const imageId = workbook.addImage({
                buffer: arrayBuffer,
                extension: "png",
            })

            sheet.addImage(imageId, {
                tl: { col: flags[i].startCol + 3, row: flags[i].startRow  },
                ext: sizeFlag, // px 단위 크기 지정
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

        <button onClick={handleExport}>엑셀로 내보내기</button><br/>

        <div>
            가로 사이즈 변경 : <br/>
            여백 : <input type="number"
                         value={widthColumns.space}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                space: Number(e.target.value)
                            })
                         }}/><br/>
            가로0 : <input type="number"
                         value={widthColumns.colums0}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums0: Number(e.target.value)
                            })
                         }}/><br/>
            가로1 : <input type="number"
                         value={widthColumns.colums1}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums1: Number(e.target.value)
                            })
                         }}/><br/>
            가로2 : <input type="number"
                         value={widthColumns.colums2}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums2: Number(e.target.value)
                            })
                         }}/><br/>
            가로3 : <input type="number"
                         value={widthColumns.colums3}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums3: Number(e.target.value)
                            })
                         }}/><br/>
            가로4 : <input type="number"
                         value={widthColumns.colums4}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums4: Number(e.target.value)
                            })
                         }}/><br/>
            가로5 : <input type="number"
                         value={widthColumns.colums5}
                         onChange={(e)=>{
                            setWidthColumns({...widthColumns,
                                colums5: Number(e.target.value)
                            })
                         }}/><br/>
        </div>

        <div>
            세로 사이즈 변경 : <br/>
            세로0 : <input type="number"
                         value={heightRows.row0}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row0: Number(e.target.value)
                            })
                         }}/><br/>

            세로1 : <input type="number"
                         value={heightRows.row1}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row1: Number(e.target.value)
                            })
                         }}/><br/>

            세로2 : <input type="number"
                         value={heightRows.row2}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row2: Number(e.target.value)
                            })
                         }}/><br/>
            세로3 : <input type="number"
                         value={heightRows.row3}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row3: Number(e.target.value)
                            })
                         }}/><br/>
            세로4 : <input type="number"
                         value={heightRows.row4}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row4: Number(e.target.value)
                            })
                         }}/><br/>
            세로5 : <input type="number"
                         value={heightRows.row5}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row5: Number(e.target.value)
                            })
                         }}/><br/>
            세로6 : <input type="number"
                         value={heightRows.row6}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row6: Number(e.target.value)
                            })
                         }}/><br/>
            세로7 : <input type="number"
                         value={heightRows.row7}
                         onChange={(e)=>{
                            setHeightRows({...heightRows,
                                row7: Number(e.target.value)
                            })
                         }}/><br/>
        </div>
        <div>
            국가 사이즈 변경 : <br/>
            가로 : <input type="number"
                         value={sizeFlag.width}
                         onChange={(e)=>{
                            setSizeFlag({...sizeFlag,
                                width: Number(e.target.value)
                            })
                         }}/><br/>
            세로 : <input type="number"
                         value={sizeFlag.height}
                         onChange={(e)=>{
                            setSizeFlag({...sizeFlag,
                                height: Number(e.target.value)
                            })
                         }}/><br/>
        </div>

        <div>now wine datas : {datas.length}<br/>
        <strong>이름, 국가, 지역, 품종, 용량, 정상가(원), 판매가(원)</strong><br/>
        {datas.map((wine, index)=>{
             if(index > 0){
                return (<div key={index}>
                {wine.name}, {wine.nation}, {wine.region}, {wine.category}, {wine.volume}, {wine.normalPrice}, {wine.nowPrice}</div>)
             }else{
                return <></>
             }
        })}
        </div>
    </>
}

export default WineComponent