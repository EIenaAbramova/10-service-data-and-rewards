import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out=process.argv[2];await fs.mkdir(out,{recursive:true});
const rows={
 '07':[
  ['D001','17,99','01.07.2024','Клиника А',60000,'Партнёр А'],
  ['D001','17,99','02.07.2024','Клиника А',50000,'Партнёр А'],
  ['D002',18,'03.07.2024','Клиника А',150000,'Партнёр А'],
  ['D003',10,'04.07.2024','Клиника Б',100000,'Партнёр Б'],
  ['D004',12,'05.07.2024','Клиника Б',200000,null]],
 '08':[
  ['D001',18,'01.08.2024','Клиника А',70000,'Партнёр А'],
  ['D001',18,'02.08.2024','Клиника А',60000,'Партнёр А'],
  ['D005',8,'03.08.2024','Клиника Б',120000,'Партнёр Б'],
  ['D006',30,'04.08.2024','Клиника Б',120001,'Партнёр Б'],
  ['D007',25,'05.08.2024','Клиника А',150000,null],
  ['D008',17,'06.08.2024','Клиника А',60000,'Партнёр А'],
  ['D008',17,'07.08.2024','Клиника Б',60000,'Партнёр А']],
 '09':[
  ['D009',9,'01.09.2024','Клиника А',130000,'Партнёр Б'],
  ['D010',40,'02.09.2024','Клиника А',120000,'Партнёр А'],
  ['D011',11,'03.09.2024','Клиника Б',120000.01,'Партнёр Б'],
  ['D012',12,'04.09.2024','Клиника Б',200000,'  ']]};
for(const [month,records]of Object.entries(rows)){
 const w=Workbook.create(),s=w.worksheets.add('EKMS');
 s.getRange(`A1:F${records.length+1}`).values=[['id пациента','возраст','дата услуги','клиника услуги','компенсация, руб.','представитель'],...records];
 s.getRange('A1:F1').format={fill:'#244763',font:{bold:true,color:'#FFFFFF'}};
 s.getRange(`A1:F${records.length+1}`).format.columnWidth=23;
 await(await SpreadsheetFile.exportXlsx(w)).save(path.join(out,`EKMS_24${month}.xlsx`));
}
console.log('Synthetic sources exported');
