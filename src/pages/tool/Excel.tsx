import { useState } from 'react';
import * as XLSX from 'xlsx';

const Excel = () => {
  const [inputs, setInputs] = useState({ first: '', second: '', third: '' });
  const [tableData, setTableData] = useState([]);
  const [isComposing, setIsComposing] = useState(false); // IME 상태 관리
  const [fileName, setFileName] = useState(''); // File name state

  const handleChange = (e) => {
    const { name, value } = e.target;

    // Restrict the third input to numbers only, with a maximum length of 8
    if (name === 'third') {
      if (!/^[0-9]*$/.test(value) || value.length > 8) {
        return;
      }
    }

    setInputs({ ...inputs, [name]: value });
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter') {
      if (isComposing) {
        return; // IME 입력 중에는 Enter 키 동작을 무시
      }

      e.preventDefault();

      // Check if any input is empty
      if (!inputs.first || !inputs.second || !inputs.third) {
        const confirmInput = window.confirm(
          '비어있는 입력란이 있는데 저장하시겠습니까?',
        );
        if (!confirmInput) {
          return; // Exit if user cancels
        }
      }

      // Ensure the third input has exactly 8 characters if it has any value
      if (inputs.third && inputs.third.length !== 8) {
        alert('전화번호 입력란은 비워두거나 숫자 8개 입력해야 합니다.');
        return;
      }

      // Format the third input if it has 8 characters
      const formattedThird = inputs.third
        ? `010-${inputs.third.slice(0, 4)}-${inputs.third.slice(4)}`
        : '';

      // Add the new row to the table data
      setTableData((prevData) => [
        ...prevData,
        { first: inputs.first, second: inputs.second, third: formattedThird },
      ]);

      // Clear the inputs
      setInputs({ first: '', second: '', third: '' });
    }
  };

  const handleCompositionStart = () => {
    setIsComposing(true);
  };

  const handleCompositionEnd = () => {
    setIsComposing(false);
  };

  const handleFileNameChange = (e) => {
    setFileName(e.target.value);
  };

  const exportToExcel = () => {
    if (!fileName) return;

    // Add row numbers to the exported data
    const numberedTableData = tableData.map((row, index) => ({
      번호: index + 1,
      이름: row.first,
      주소: row.second,
      전화번호: row.third,
    }));

    const worksheet = XLSX.utils.json_to_sheet(numberedTableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    XLSX.writeFile(workbook, `${fileName}.xlsx`);

    // Clear the file name input
    setFileName('');
  };

  const importFromExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      // Update table data from imported Excel
      const formattedData = jsonData.map((row, index) => ({
        first: row['이름'] || `Row ${index + 1}`,
        second: row['주소'] || '',
        third: row['전화번호'] || '',
      }));

      setTableData(formattedData);
    };

    reader.readAsArrayBuffer(file);
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Table Example</h1>

      {/* Input Fields */}
      <div style={{ marginBottom: '20px' }}>
        <input
          type="text"
          name="first"
          value={inputs.first}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onCompositionStart={handleCompositionStart}
          onCompositionEnd={handleCompositionEnd}
          placeholder="이름"
          style={{ marginRight: '10px', padding: '5px' }}
        />
        <input
          type="text"
          name="second"
          value={inputs.second}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onCompositionStart={handleCompositionStart}
          onCompositionEnd={handleCompositionEnd}
          placeholder="주소"
          style={{ marginRight: '10px', padding: '5px' }}
        />
        <input
          type="text"
          name="third"
          value={inputs.third}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onCompositionStart={handleCompositionStart}
          onCompositionEnd={handleCompositionEnd}
          placeholder="전화번호"
          style={{ padding: '5px' }}
        />
      </div>

      {/* File Name Input */}
      <div style={{ marginBottom: '20px', textAlign: 'right' }}>
        <input
          type="text"
          value={fileName}
          onChange={handleFileNameChange}
          placeholder="파일명을 입력하세요"
          style={{ marginRight: '10px', padding: '5px' }}
        />
        <button
          onClick={exportToExcel}
          disabled={!fileName}
          style={{
            padding: '5px 10px',
            cursor: fileName ? 'pointer' : 'not-allowed',
          }}
        >
          엑셀로 내보내기
        </button>
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={importFromExcel}
          style={{ marginLeft: '10px' }}
        />
      </div>

      {/* Table */}
      <table border="1" style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            <th>#</th>
            <th>이름</th>
            <th>주소</th>
            <th>전화번호</th>
          </tr>
        </thead>
        <tbody>
          {tableData.map((row, index) => (
            <tr key={index}>
              <td style={{ textAlign: 'center' }}>{index + 1}</td>
              <td>{row.first}</td>
              <td>{row.second}</td>
              <td>{row.third}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Excel;
