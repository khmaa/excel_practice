import { useState } from 'react';
import DaumPostcode from 'react-daum-postcode';
import * as XLSX from 'xlsx';

const Excel = () => {
  const [inputs, setInputs] = useState({ first: '', second: '', third: '' });
  const [tableData, setTableData] = useState([]);
  const [isComposing, setIsComposing] = useState(false); // IME 상태 관리
  const [fileName, setFileName] = useState(''); // File name state

  const [openPostcode, setOpenPostcode] = useState(false);

  const [calendarLocation, setCalendarLocation] = useState('');
  const locations = { calendarLocation: calendarLocation };

  const handleChange = (e) => {
    const { name, value } = e.target;

    // Restrict the third input to numbers only, with a maximum length of 8
    if (name === 'third') {
      if (!/^[0-9]*$/.test(value) || value.length > 8) {
        return;
      }
    }

    // Trim spaces for the first input
    const trimmedValue = name === 'first' ? value.replace(/\s+/g, '') : value;

    setInputs({ ...inputs, [name]: trimmedValue });
  };

  const handleOpenPostCode = () => {
    setOpenPostcode((current) => !current);
  };

  const handleSelectAddress = (data) => {
    setCalendarLocation(data.address);
    setOpenPostcode(false);
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter') {
      if (isComposing) {
        return; // IME 입력 중에는 Enter 키 동작을 무시
      }

      e.preventDefault();

      // Ensure the first input is not empty
      if (!inputs.first) {
        alert('이름이 비어있습니다');
        return;
      }

      // Check if any other input is empty
      if (!calendarLocation || !inputs.third) {
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
        {
          first: inputs.first,
          second: calendarLocation,
          third: formattedThird,
        },
      ]);

      // Clear the inputs
      setInputs({ first: '', second: '', third: '' });
      setCalendarLocation('');
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
          value={calendarLocation}
          // onChange={handleChange}
          // onKeyDown={handleKeyDown}
          // onCompositionStart={handleCompositionStart}
          // onCompositionEnd={handleCompositionEnd}
          onClick={handleOpenPostCode}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault(); // 기본 엔터 동작 방지
              handleOpenPostCode(); // PostCode 열기 동작
            }
          }}
          placeholder="주소"
          style={{
            marginRight: '10px',
            padding: '5px',
            width: calendarLocation ? '400px' : 'auto',
          }}
        />

        {!calendarLocation && (
          <button
            type="button"
            onClick={handleOpenPostCode}
            style={{ marginRight: '20px' }}
          >
            {calendarLocation ? calendarLocation : '장소를 검색해주세요'}
          </button>
        )}
        {openPostcode && (
          <DaumPostcode
            onComplete={handleSelectAddress} // 값을 선택할 경우 실행되는 이벤트
            autoClose={false} // 값을 선택할 경우 사용되는 DOM을 제거하여 자동 닫힘 설정
          />
        )}

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
          disabled={!fileName || tableData.length < 1}
          style={{
            padding: '5px 10px',
            cursor: fileName ? 'pointer' : 'not-allowed',
          }}
        >
          엑셀로 내보내기
        </button>
        <label
          htmlFor="file-upload"
          style={{
            marginLeft: '10px',
            padding: '5px 10px',
            backgroundColor: '#007BFF',
            color: '#fff',
            cursor: 'pointer',
            borderRadius: '4px',
            display: 'inline-block',
          }}
        >
          엑셀 파일 불러오기
        </label>
        <input
          id="file-upload"
          type="file"
          accept=".xlsx, .xls"
          onChange={importFromExcel}
          style={{ display: 'none' }}
        />
        {/* <input
          type="file"
          accept=".xlsx, .xls"
          onChange={importFromExcel}
          style={{ marginLeft: '10px' }}
        /> */}
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
