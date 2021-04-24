import { ReactExcel, readFile, generateObjects } from '@ramonak/react-excel';
import React, { useState } from 'react';
import { signatureVerify } from '@polkadot/util-crypto'
import XLSX from 'xlsx'

import './index.css'

function exportExcel(headers, data, fileName = 'verify.xlsx') {
  const _headers = headers
      .map((item, i) => Object.assign({}, { key: item.key, title: item.title, position: String.fromCharCode(65 + i) + 1 }))
      .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { key: next.key, v: next.title } }), {});

  const _data = data
      .map((item, i) => headers.map((key, j) => Object.assign({}, { content: item[key.key], position: String.fromCharCode(65 + j) + (i + 2) })))
      .reduce((prev, next) => prev.concat(next))
      .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { v: next.content } }), {});

  const output = Object.assign({}, _headers, _data);
  const outputPos = Object.keys(output);
  const ref = `${outputPos[0]}:${outputPos[outputPos.length - 1]}`;

  const wb = {
      SheetNames: ['mySheet'],
      Sheets: {
          mySheet: Object.assign(
              {},
              output,
              {
                  '!ref': ref,
                  '!cols': [{ wpx: 45 }, { wpx: 100 }, { wpx: 200 }, { wpx: 80 }, { wpx: 150 }, { wpx: 100 }, { wpx: 300 }, { wpx: 300 }],
              },
          ),
      },
  };

  XLSX.writeFile(wb, fileName);
}
 
const App = () => {
  const [initialData, setInitialData] = useState(undefined);
  const [currentSheet, setCurrentSheet] = useState({});
 
  const handleUpload = (event) => {
    const file = event.target.files[0];

    readFile(file)
      .then((readedData) => setInitialData(readedData))
      .catch((error) => console.error(error));
  };
 
  const save = () => {
    const sheet = generateObjects(currentSheet);
    const validatedSheet = sheet.map(item => {
      if (!item.Message || !item.Signature || !item.Address) {
        return{...item, Result: '无效验证'}
      }

      let verifyResult;

      try {
        verifyResult = signatureVerify(String(item.Message), String(item.Signature), item.Address).isValid 
      } catch {
        verifyResult = false
      }

      return {
        ...item, 
        Result: verifyResult ? '验证通过' : '验证失败'
      }
    })

    const initColumn = [
      {
        title: 'Message',
        dataIndex: 'Message',
        key: 'Message',
      }, 
      {
        title: 'Signature',
        dataIndex: 'Signature',
        key: 'Signature',
      }, 
      {
        title: 'Address',
        dataIndex: 'Address',
        key: 'Address',
      },
      {
        title: 'Result',
        dataIndex: 'Result',
        key: 'Result',
      }
    ];
    console.log(validatedSheet)
    exportExcel(initColumn, validatedSheet, 'verify.xlsx')
  };
 
  return (
    <>
      <div className='input'>
        <span>表格类型为xlsx, 表头为Message, Signature, Address</span>
        <input
          type='file'
          accept='.xlsx'
          onChange={handleUpload}
        />
        <button onClick={save}>
          Export Validated Sheet
        </button>
      </div>
      <ReactExcel
        initialData={initialData}
        onSheetUpdate={(currentSheet) => setCurrentSheet(currentSheet)}
        activeSheetClassName='active-sheet'
        reactExcelClassName='react-excel'
      />
    </>
  );
}

export default App;
