import { useEffect, useRef, useState } from 'react';
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
  const [selectedRowIndex, setSelectedRowIndex] = useState(null); // 선택된 행의 인덱스

  const [isDisabledTempButton, setIsDisabledTempButton] = useState(false);

  const inputNameRef = useRef(null);
  const inputPhoneRef = useRef(null);

  const handleChange = (e) => {
    const { name, value } = e.target;

    // Restrict the third input to numbers only, with a maximum length of 8
    if (name === 'third') {
      const sanitizedValue = value.replace(/-/g, ''); // Remove dashes
      if (!/^[0-9]*$/.test(sanitizedValue) || sanitizedValue.length > 8) {
        return;
      }
      setInputs({ ...inputs, [name]: sanitizedValue });
    } else {
      // Trim spaces for the first input
      const trimmedValue = name === 'first' ? value.replace(/\s+/g, '') : value;
      setInputs({ ...inputs, [name]: trimmedValue });
    }
  };

  const handleOpenPostCode = () => {
    setOpenPostcode((current) => !current);
    setTimeout(() => {
      window.scrollTo({
        top: document.body.scrollHeight,
        behavior: 'smooth',
      });
    });
  };

  const handleSelectAddress = (data) => {
    if (data.query.slice(-1) === '동') {
      let searchText = '';
      if (data.query.includes(' ')) {
        searchText = data.query.slice(data.query.lastIndexOf(' ') + 1);
      } else {
        searchText = data.query;
      }
      setCalendarLocation(
        data.sido + ' ' + data.sigungu + ' ' + searchText,
        // (data.hname ? data.hname : data.query),
      );
    } else if (data.query.slice(-1) === '구') {
      setCalendarLocation(data.sido + ' ' + data.sigungu);
    } else if (data.query.slice(-1) === '시') {
      if (data.query === '서울시') setCalendarLocation(data.sido);
      else setCalendarLocation(data.sido + ' ' + data.query);
    } else {
      setCalendarLocation(data.sido + ' ' + data.sigungu + ' ' + data.roadname);
    }
    setOpenPostcode(false);
  };

  const checkForDuplicates = () => {
    const duplicates = []; // 중복 그룹 저장

    // 이미 처리한 행을 기록
    const processedIndices = new Set();

    tableData.forEach((row, index) => {
      if (processedIndices.has(index)) return; // 이미 처리된 인덱스는 건너뜀

      const duplicateIndices = tableData
        .map((item, i) =>
          i !== index && row.first === item.first && row.third === item.third
            ? i
            : null,
        )
        .filter((i) => i !== null);

      if (duplicateIndices.length > 0) {
        const group = [index + 1, ...duplicateIndices.map((i) => i + 1)]; // 1-based index
        duplicates.push(group);
        group.forEach((i) => processedIndices.add(i - 1)); // 처리된 인덱스로 기록
      }
    });

    if (duplicates.length > 0) {
      alert(
        `중복된 값들: ${duplicates.map((group) => `[${group.join(', ')}]`).join(' , ')}`,
      );
    } else {
      alert('중복된 값이 없습니다.');
    }
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

      const koreanRegex = /^[가-힣]+$/; // 완성된 한글만 허용
      if (!koreanRegex.test(inputs.first)) {
        alert('이름 입력란에 완성되지 않은 한글이 포함되어 있습니다.');
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

      // Add or update the row in the table data
      if (selectedRowIndex !== null) {
        // Update the existing row
        setTableData((prevData) =>
          prevData.map((row, index) =>
            index === selectedRowIndex
              ? {
                  first: inputs.first,
                  second: calendarLocation,
                  third: formattedThird,
                }
              : row,
          ),
        );
        setSelectedRowIndex(null); // Clear the selected row
      } else {
        // Add a new row
        setTableData((prevData) => [
          ...prevData,
          {
            first: inputs.first,
            second: calendarLocation,
            third: formattedThird,
          },
        ]);
      }

      // Clear the inputs
      setInputs({ first: '', second: '', third: '' });
      setCalendarLocation('');
      if (inputNameRef && inputNameRef.current) inputNameRef.current.focus();

      setTimeout(() => {
        window.scrollTo({
          top: document.body.scrollHeight,
          behavior: 'smooth',
        });
      });
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

  const appendFromExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      // Append new rows to the existing table data
      const formattedData = jsonData.map((row, index) => ({
        first: row['이름'] || `Row ${index + 1}`,
        second: row['주소'] || '',
        third: row['전화번호'] || '',
      }));

      setTableData((prevData) => [...prevData, ...formattedData]);
    };

    reader.readAsArrayBuffer(file);
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

      // Set the file name input to the uploaded file's name
      setFileName(file.name.replace(/\.[^/.]+$/, '')); // Remove the file extension
    };

    reader.readAsArrayBuffer(file);
  };

  const handleRowClick = (index) => {
    const row = tableData[index];
    setInputs({
      first: row.first,
      second: '',
      third: row.third.replace(/-/g, '').replace(/^010/, ''),
    });
    setCalendarLocation(row.second);
    setSelectedRowIndex(index);
  };

  const handleSaveToLocalStorage = () => {
    if (confirm('임시 저장하시겠습니까?')) {
      const dataToSave = {
        inputs,
        tableData,
        calendarLocation,
        fileName,
      };
      localStorage.setItem('excelAppData', JSON.stringify(dataToSave));
      alert('현재 화면 정보가 임시저장되었습니다.');
    }
  };

  const handleLoadFromLocalStorage = () => {
    const savedData = localStorage.getItem('excelAppData');
    if (savedData) {
      const parsedData = JSON.parse(savedData);

      setInputs(parsedData.inputs || { first: '', second: '', third: '' });
      setTableData(parsedData.tableData || []);
      setCalendarLocation(parsedData.calendarLocation || '');
      setFileName(parsedData.fileName || '');

      // 성공적으로 불러온 후 localStorage에서 데이터 삭제
      localStorage.removeItem('excelAppData');
      alert(
        '임시 저장된 데이터를 불러왔습니다. 불러온 데이터는 삭제되었습니다.',
      );
    } else {
      alert('불러올 데이터가 없습니다.');
    }
  };

  const resetInputs = () => {
    if (confirm('입력창을 초기화 하시겠습니까?')) {
      setInputs({ first: '', second: '', third: '' });
      setCalendarLocation('');
      setSelectedRowIndex(null);
    }
  };

  useEffect(() => {
    if (calendarLocation) {
      if (inputPhoneRef && inputPhoneRef.current) inputPhoneRef.current.focus();
    }
  }, [calendarLocation]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault(); // Prevent the default browser save action
        handleSaveToLocalStorage(); // Call the save function
      }
    };

    window.addEventListener('keydown', handleKeyDown);

    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [inputs, tableData, calendarLocation, fileName]);

  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (tableData.length > 0) {
        const confirmationMessage =
          '저장되지 않은 데이터가 있습니다. 진행하시겠습니까?';
        e.preventDefault();
        e.returnValue = confirmationMessage; // 일부 브라우저에서 동작
        return confirmationMessage;
      }
    };

    window.addEventListener('beforeunload', handleBeforeUnload);

    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [tableData]);

  useEffect(() => {
    if (typeof window !== 'undefined' && localStorage.getItem('excelAppData')) {
      setIsDisabledTempButton(true);
    } else {
      setIsDisabledTempButton(false);
    }
  }, []);

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Table Example</h1>
      <div>
        <button
          onClick={checkForDuplicates}
          style={{
            position: 'absolute',
            top: '55px',
            right: '380px',
            padding: '10px 15px',
            backgroundColor: '#FF4500',
            color: '#fff',
            border: 'none',
            borderRadius: '5px',
            cursor: 'pointer',
          }}
        >
          중복 확인
        </button>
        <button
          onClick={handleLoadFromLocalStorage}
          disabled={tableData.length > 0}
          style={{
            position: 'absolute',
            top: '55px',
            right: '130px',
            padding: '10px 15px',
            backgroundColor: isDisabledTempButton ? '#4CAF50' : '#ccc',
            color: '#fff',
            border: 'none',
            borderRadius: '5px',
            cursor: isDisabledTempButton ? 'pointer' : 'not-allowed',
          }}
        >
          불러오기
        </button>

        <button
          onClick={handleSaveToLocalStorage}
          style={{
            position: 'absolute',
            top: '55px',
            right: '10px',
            padding: '10px 15px',
            backgroundColor: '#FFA500',
            color: '#fff',
            border: 'none',
            borderRadius: '5px',
            cursor: 'pointer',
          }}
        >
          임시저장
        </button>
      </div>
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
            <tr
              key={index}
              onDoubleClick={() => handleRowClick(index)}
              style={{
                backgroundColor:
                  selectedRowIndex === index ? '#d4e6fb' : 'white',
                cursor: 'pointer',
              }}
            >
              <td style={{ textAlign: 'center' }}>{index + 1}</td>
              <td>{row.first}</td>
              <td>{row.second}</td>
              <td>{row.third}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ marginTop: '30px', marginBottom: '30px' }}>
        <div style={{ marginBottom: '20px' }}>
          <input
            type="text"
            name="first"
            ref={inputNameRef}
            value={inputs.first}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            onCompositionStart={handleCompositionStart}
            onCompositionEnd={handleCompositionEnd}
            placeholder="이름"
            style={{ marginRight: '10px', padding: '5px' }}
            maxLength={5}
          />
          <input
            type="text"
            name="second"
            value={calendarLocation}
            onFocus={(e) => {
              e.preventDefault();
              handleOpenPostCode();
            }}
            onClick={handleOpenPostCode}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                handleOpenPostCode();
              }
            }}
            placeholder="주소"
            style={{
              marginRight: '10px',
              padding: '5px',
              width: calendarLocation ? '200px' : 'auto',
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
            <DaumPostcode onComplete={handleSelectAddress} autoClose={false} />
          )}

          <input
            type="text"
            name="third"
            ref={inputPhoneRef}
            value={inputs.third}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            onCompositionStart={handleCompositionStart}
            onCompositionEnd={handleCompositionEnd}
            placeholder="전화번호"
            style={{ padding: '5px' }}
          />
        </div>

        <div style={{ marginBottom: '20px', textAlign: 'right' }}>
          <label
            htmlFor="file-append"
            style={{
              marginRight: '10px',
              padding: '5px 10px',
              backgroundColor: '#28A745',
              color: '#fff',
              cursor: 'pointer',
              borderRadius: '4px',
              display: 'inline-block',
            }}
          >
            엑셀 이어 붙이기
          </label>
          <input
            id="file-append"
            type="file"
            accept=".xlsx, .xls"
            onChange={appendFromExcel}
            style={{ display: 'none' }}
          />
          <button
            onClick={resetInputs}
            style={{
              marginRight: '10px',
              padding: '5px 10px',
              backgroundColor: '#FF5733',
              color: '#fff',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
            }}
          >
            입력창 초기화
          </button>
          <input
            type="text"
            value={fileName}
            onChange={handleFileNameChange}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                exportToExcel();
              }
            }}
            placeholder="파일명을 입력하세요"
            style={{ marginRight: '10px', padding: '5px' }}
            maxLength={30}
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
        </div>
      </div>
    </div>
  );
};

export default Excel;
