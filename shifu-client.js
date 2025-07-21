let urlAPI;
const MAIN_SHEET_NAME = 'MÀN HÌNH CHÍNH'; // Tên sheet chính
const CUSTOMER_SHEET_NAME = 'KHÁCH ĐÃ LƯU';
const NAME_CELL       = 'C3'; // Ô chứa họ tên
const PROFESSION_CELL = 'C4'; // Ô chứa nghề nghiệp
const AGE_CELL        = 'C7'; // Ô chứa tuổi
const GENDER_CELL     = 'E7'; // Ô chứa giới tính
const CCCD_CELL       = 'C5'; // Ô chứa CCCD
const PHONE_CELL      = 'C6'; // Ô chứa SĐT

const LUANSIM_CELL    = 'B29'; // Ô để in kết quả ra Luận Sim
const BIENHOA_CELL    = 'R23'; // Ô để in kết quả ra Biến Hóa
const HOAHUNG_CELL    = 'O23:P29'; // Ô in ra kết quả Hóa Hung với 0

const API_LEFT_NUM_CELL = 'L25'; // Ô in ra số lượt sử dụng AI còn lại


const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
const MAX_NUM_AI      = 30 ;

const wsMain = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
const wsCustomer = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMER_SHEET_NAME);


//Lấy setting
function getMasterSettingWithCache(rowNumber, columnLetter) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `urlAPI_${rowNumber}_${columnLetter}`;
  const cachedValue = cache.get(cacheKey);

  if (cachedValue) {
    return cachedValue;
  }

  // Nếu chưa có trong cache → lấy từ file Master
  const IDMaster = "1UvQOYMpF2FqFws4brrHlbH16-QItBpeTNZqV251YYwg"; 
  const master = SpreadsheetApp.openById(IDMaster);
  const sheet = master.getSheetByName("SETTING");

  if (!sheet) {
    throw new Error("Không tìm thấy sheet 'SETTING' trong file master.");
  }

  if (!rowNumber || typeof rowNumber !== 'number' || rowNumber < 1) {
    throw new Error("Giá trị rowNumber không hợp lệ. Phải là số nguyên dương.");
  }

  if (!columnLetter || typeof columnLetter !== 'string' || !/^[A-Z]+$/.test(columnLetter)) {
    throw new Error("Giá trị columnLetter không hợp lệ. Phải là chữ cái in hoa từ A-Z.");
  }

  const cellAddress = columnLetter + rowNumber;
  const value = sheet.getRange(cellAddress).getValue();

  if (!value || String(value).trim() === "") {
    throw new Error(`Ô ${cellAddress} rỗng hoặc không có giá trị hợp lệ.`);
  }

  // Lưu vào cache trong 6 giờ (21600 giây)
  cache.put(cacheKey, value, 21600);

  // Lưu vào ScriptProperties nếu cần dùng lại
  PropertiesService.getScriptProperties().setProperty("urlAPI", value);

  return value;
}


function phanTichSoDienThoai_LanLuot(sdt) {
  var s1 = tachHoaHung(sdt);
  var s2 = tachKichPhat(s1);
  var s3 = tachPhucVi(s2);
  return s3;
}

function tachHoaHung(sdt) {
  var arr = ["00000", "0000", "000", "00", "0"];
  return boCacBoSo(sdt, arr);
}

function tachKichPhat(sdt) {
  var arr = ["55555", "5555", "555", "55", "5"];
  return boCacBoSo(sdt, arr);
}

function tachPhucVi(sdt) {
  var patterns = [
    "11111:1", "1111:1", "111:1", "11:1",
    "22222:2", "2222:2", "222:2", "22:2",
    "33333:3", "3333:3", "333:3", "33:3",
    "44444:4", "4444:4", "444:4", "44:4",
    "66666:6", "6666:6", "666:6", "66:6",
    "77777:6", "7777:6", "777:6", "77:6",
    "88888:8", "8888:8", "888:8", "88:8",
    "99999:9", "9999:9", "999:9", "99:9"
  ];

  var i = 0;
  var result = "";

  while (i < sdt.length) {
    var matched = false;
    for (var j = 0; j < patterns.length; j++) {
      var parts = patterns[j].split(":");
      var key = parts[0];
      var rep = parts[1];
      if (sdt.substring(i, i + key.length) === key) {
        result += rep;
        i += key.length;
        matched = true;
        break;
      }
    }
    if (!matched) {
      result += sdt[i];
      i++;
    }
  }
  return result;
}

function boCacBoSo(sdt, arr) {
  var i = 0;
  var result = "";

  while (i < sdt.length) {
    var matched = false;
    for (var j = 0; j < arr.length; j++) {
      var sub = arr[j];
      if (sdt.substring(i, i + sub.length) === sub) {
        i += sub.length;
        matched = true;
        break;
      }
    }
    if (!matched) {
      result += sdt[i];
      i++;
    }
  }
  return result;
}
function saveCustomer() {
  
  const name = cleanText(wsMain.getRange(NAME_CELL).getCell(1,1).getDisplayValue());
  const profession = cleanText(wsMain.getRange(PROFESSION_CELL).getCell(1,1).getDisplayValue());
  const age = cleanText(wsMain.getRange(AGE_CELL).getCell(1,1).getDisplayValue());
  const gender = cleanText(wsMain.getRange(GENDER_CELL).getCell(1,1).getDisplayValue());
  const cccd = cleanText(wsMain.getRange(CCCD_CELL).getCell(1,1).getDisplayValue());
  const phone = cleanText(wsMain.getRange(PHONE_CELL).getCell(1,1).getDisplayValue());
  
  if (!isValidPhoneNumber(phone)) {
    SpreadsheetApp.getActive().toast("Số điện thoại không đúng định dạng", "Thầy đang nói", 5);
    return;
  }

  if (!isValidCCCDNumber(cccd)) {
    SpreadsheetApp.getActive().toast("Số CCCD không đúng định dạng", "Thầy đang nói", 5);    return;
  }

  const newRow = wsCustomer.getLastRow() + 1;
  wsCustomer.getRange(newRow, 1).setValue(new Date());
  wsCustomer.getRange(newRow, 2).setValue(name);
  wsCustomer.getRange(newRow, 3).setValue(phone);
  wsCustomer.getRange(newRow, 4).setValue(cccd);
  wsCustomer.getRange(newRow, 5).setValue(age);
  wsCustomer.getRange(newRow, 6).setValue(profession);
  wsCustomer.getRange(newRow, 7).setValue(gender);

  
  wsMain.getRange(NAME_CELL).getCell(1,1).clearContent(); // === Ho ten
  wsMain.getRange(PROFESSION_CELL).getCell(1,1).clearContent(); // === Nghề
  wsMain.getRange(AGE_CELL).getCell(1,1).clearContent(); // === năm sinh
  wsMain.getRange(CCCD_CELL).getCell(1,1).clearContent(); // === CCCD
  wsMain.getRange(PHONE_CELL).getCell(1,1).clearContent(); // === SĐT
  SpreadsheetApp.getActive().toast("ĐÃ LƯU KHÁCH HÀNG THÀNH CÔNG, XIN MỜI XEM CHO NGƯỜI MỚI", "Thầy đang nói", 5);
  

}

function isValidPhoneNumber(sdt) {
  return /^0\d{9}$/.test(sdt);
}

function isValidCCCDNumber(cccd) {
  return /^\d{6}$|^\d{12}$/.test(cccd);
}


function cleanText(txt) {
  return (txt + "").replace(/[\n\r]/g, "").trim();
}

function normalizePhoneNumber(inputText) {
  let cleaned = inputText.replace(/@/g, "0").replace(/[.\s]/g, "");
  if (cleaned.startsWith("+84")) {
    cleaned = "0" + cleaned.substring(3);
  }
  return cleaned.trim();
}

function xuLySoDienThoai(sdt) {
  // Bước 1: Bỏ số 0 và 5
  let cleaned = String(sdt).replace(/0|5/g, '');

  // Bước 2: Loại bỏ số trùng liên tiếp
  let result = '';
  for (let i = 0; i < cleaned.length; i++) {
    if (i === 0 || cleaned[i] !== cleaned[i - 1]) {
      result += cleaned[i];
    }
  }

  return result;
}

function xuLyGiuMotSo0(sdt) {
  let text = String(sdt);

  // Giữ lại 1 số 0 nếu có nhiều số 0 liên tiếp
  let giuMotSo0 = text.replace(/0{2,}/g, '0');

  // Bỏ số 5
  let boSo5 = giuMotSo0.replace(/5/g, '');

  // Bỏ số lặp liên tiếp
  let result = '';
  for (let i = 0; i < boSo5.length; i++) {
    if (i === 0 || boSo5[i] !== boSo5[i - 1]) {
      result += boSo5[i];
    }
  }

  return result;
}

function boPhucVi(sdt) {
  // Bỏ ký tự sau nếu gặp 0 và hai bên là giống nhau (như 808 → 80)
  let result = '';
  for (let i = 0; i < sdt.length; i++) {
    if (
      sdt[i] === '0' &&
      i > 0 &&
      i < sdt.length - 1 &&
      sdt[i - 1] === sdt[i + 1]
    ) {
      result += '0'; // vẫn giữ 0
      i++; // bỏ qua ký tự sau 0
    } else {
      result += sdt[i];
    }
  }
  return result;
}
function xuLyCCCD() {
  
  let cccd = wsMain.getRange(CCCD_CELL).getValue().toString().trim();

  // Chỉ xử lý nếu là 6 hoặc 12 chữ số
  if (!/^\d{6}$|^\d{12}$/.test(cccd)) {
    wsMain.getRange("R2").setValue("000000000000");
    return;
  }

  // Tách thành mảng ký tự
  let chars = cccd.split("");
  let changed;

  // Biến hóa các số: nếu gặp 0 hoặc 5 và trước đó không phải 0 thì thay thế
  do {
    changed = false;
    for (let i = 1; i < chars.length; i++) {
      if ((chars[i] === '0' || chars[i] === '5') && chars[i - 1] !== '0') {
        chars[i] = chars[i - 1];
        changed = true;
      }
    }
  } while (changed);

  const result = chars.join("");
  wsMain.getRange("R2").setValue(result);
  SpreadsheetApp.getActive().toast("ĐÃ TÁCH TỪ TRƯỜNG CCCD THÀNH CÔNG", "Thầy đang nói", 5);
}




function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();


  // Kiểm tra nếu chỉnh sửa tại ô C5
  if (row === 5 && col >= 3 && col <= 5) {
    xuLyCCCD(); // gọi hàm xử lý chính
  }
  
}

/**
 * Gửi một prompt đến  API và in kết quả vào ô B26 với định dạng yêu cầu.
 */
function LuanSoThongMinh() {
  /*
  const currentLeftUsageAI = wsMain.getRange(API_LEFT_NUM_CELL).getValue();
  if(currentLeftUsageAI == 0){
    const msg = 'Bạn đã hết lượt sử dụng tính năng Luận sim AI. Vui lòng đăng ký thêm lượt dùng AI hoặc chờ đến 0h00 ngày mai để tiếp tục sử dụng';
    wsMain.getRange(LUANSIM_CELL).setValue(msg);
    SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
    return;
  }
  */
  //Logger.log(currentLeftUsageAI);
  const DROPDOWN_STYLE_CELL = 'G69'; // Ô chứa menu thả xuống phong cách luận
  const DROPDOWN_ORDER_CELL = 'G70'; // Ô chứa menu thả xuống thứ tự luận

  if (!wsMain) {
    SpreadsheetApp.getActive().toast("Không tìm thấy sheet với tên '" + MAIN_SHEET_NAME + "'. Vui lòng kiểm tra lại tên wsMain.", "Thầy đang nói", 5);
    
    return;
  }

  
  wsMain.getRange(LUANSIM_CELL).setValue('ĐỢI TÍ THẦY ĐANG LUẬN...');
  SpreadsheetApp.flush(); // Đảm bảo thay đổi được hiển thị ngay lập tức

  const inputData = KetQuaLuanThuCong(); 
  //Logger.log(inputData);
  
  if (!inputData || (typeof inputData !== 'string') ||
      (!inputData.toLowerCase().includes('ưu điểm')) ||
      (!inputData.toLowerCase().includes('nhược điểm'))) {
    const msg = "Bạn chưa thực hiện Luận giải sim hoặc dữ liệu vào không hợp lệ (thiếu 'ưu điểm'/'nhược điểm').";
    wsMain.getRange(LUANSIM_CELL).setValue(msg); 
    SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
    return; // Dừng hàm nếu điều kiện không được đáp ứng
  }
  const profession = wsMain.getRange(PROFESSION_CELL).getValue();
  const age = wsMain.getRange(AGE_CELL).getValue();
  const gender = wsMain.getRange(GENDER_CELL).getValue();
  let personalizationInfo = "";
      if (profession) personalizationInfo += `Họ là là một ${profession}. `;
      if (gender) personalizationInfo += `Giới tính ${gender}. `;
      if (age) personalizationInfo += `Ở độ tuổi ${age}. `;

      if (!personalizationInfo) {
        const msg = "Vui lòng nhập thông tin nghề nghiệp, tuổi, giới tính để cá nhân hóa.";
        SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
        return;
      }
  const ddrStyle = wsMain.getRange(DROPDOWN_STYLE_CELL).getValue();
  const ddrOrder = wsMain.getRange(DROPDOWN_ORDER_CELL).getValue();
  
  let generatedEssay = "";

  try {
    // Gọi File B với action "GenerateEssay" để tạo toàn bộ luận giải AI
    const payloadForEssayGeneration = {
      action: "GetPhongCachAIAPI", 
      ddrStyle: ddrStyle,
      ddrOrder: ddrOrder,
      inputData: inputData,
      personInfo: personalizationInfo,
      fileId: fileId
    };

    const optionsForEssayGeneration = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payloadForEssayGeneration)
    };

    
    const urlAPI = PropertiesService.getScriptProperties().getProperty("urlAPI");
    const response = UrlFetchApp.fetch(urlAPI, optionsForEssayGeneration);
    const responseCode = response.getResponseCode();
    ketQua = JSON.parse(response.getContentText());

    if (responseCode >= 400) {
      throw new Error(`Lỗi khi tạo luận giải AI từ File B (${responseCode}): ${ketQua}`);
    }

    wsMain.getRange(LUANSIM_CELL).setValue(ketQua.textResult);
    wsMain.getRange(API_LEFT_NUM_CELL).setValue(ketQua.maxUsageAI - ketQua.curentUsageAI)
    SpreadsheetApp.getActive().toast("LUẬN THÀNH CÔNG", "Thầy đang nói", 5);

  } catch (err) {
    const errorMsg = `Lỗi khi tạo Luận giải thông minh: ${err.message}`;
    wsMain.getRange(LUANSIM_CELL).setValue(errorMsg);
    SpreadsheetApp.getActive().toast(errorMsg, "Thầy đang nói", 5);
  }
}
function LuanSoThuCong() {
  
  wsMain.getRange(LUANSIM_CELL).setValue("ĐANG PHÂN TÍCH CÁC TỪ TRƯỜNG...");
  const KetQua = JSON.parse(KetQuaLuanThuCong());
  wsMain.getRange(LUANSIM_CELL).setValue(KetQua.textResult);
  wsMain.getRange(API_LEFT_NUM_CELL).setValue(KetQua.maxUsageAI - KetQua.curentUsageAI);
  SpreadsheetApp.getActive().toast("LUẬN THÀNH CÔNG", "Thầy đang nói", 5);
  
}
function KetQuaLuanThuCong(){
  const sdt = wsMain.getRange(PHONE_CELL).getDisplayValue().trim();
  const cache = CacheService.getDocumentCache();
  const cacheKey = 'KetQuaLuanThuCong_' + sdt;

  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    Logger.log("✅ Đã có cache với key: " + cacheKey);
    return cachedData; 
  }
  

  // Lấy dữ liệu từ Thong Tin Sim
  const data = wsMain.getRange("N8:R19").getValues();
  // Lấy dữ liệu từ Q23 cho phần cảnh báo đặc biệt
  const specialWarningContent = wsMain.getRange(BIENHOA_CELL).getValue();
  
  const hoaHungData = wsMain.getRange("N23:P29").getValues();
  
  
  const payloadForLuanThuCong = {
    action: "LuanThuCongAPI",
    data: data,
    specialWarningContent: specialWarningContent,
    hoaHungData: hoaHungData,
    fileId: fileId
  };

  

  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payloadForLuanThuCong)
    };
    const urlAPI = PropertiesService.getScriptProperties().getProperty("urlAPI");
    //Logger.log(urlAPI);
    
    const response = UrlFetchApp.fetch(urlAPI, options);
    const responseCode = response.getResponseCode();
    const ketQua = response.getContentText();
    
    // Kiểm tra mã phản hồi HTTP
    if (responseCode >= 200 && responseCode < 300) {
      
      cache.put(cacheKey, ketQua, 300); // Cache theo sđt trong 5 phút
      return ketQua;

    } else {
      // Trả về lỗi nếu có vấn đề về HTTP
      return "Lỗi từ API (" + responseCode + "): " + ketQua;
    }
  } catch (err) {
    //Logger.log("Error in KetQuaLuanThuCong: " + err.message);
    return "Lỗi khi gọi API: " + err.message;
  }
  
}
function KhenSim() {
  

  const CHECKBOX_START_ROW = 71; // Hàng bắt đầu của các ô kiểm (G71)
  const CHECKBOX_END_ROW = 79; // Hàng kết thúc của các ô kiểm (G79)
  const DEMAND_COLUMN_INDEX = (8); // Cột H (index 8) cho dữ liệu nhu cầu
  const HIGHLIGHT_CELL_START = 'Q23'; // Ô bắt đầu dữ liệu điểm nổi bật của sim
  const HIGHLIGHT_CELL_END = 'Q29'; // Ô kết thúc dữ liệu điểm nổi bật của sim

  if (!wsMain) {
    //Logger.log("Không tìm thấy sheet với tên '" + MAIN_SHEET_NAME + "'");
    SpreadsheetApp.getActive().toast("Không tìm thấy sheet với tên '" + MAIN_SHEET_NAME + "'", "Thầy đang nói", 5);
    
    return;
  }
  /*
  const currentLeftUsageAI = wsMain.getRange(API_LEFT_NUM_CELL).getValue();
  if(currentLeftUsageAI == 0){
    const msg = 'Bạn đã hết lượt sử dụng tính năng Khen sim AI. Vui lòng đăng ký thêm lượt dùng AI hoặc chờ đến 0h00 ngày mai để tiếp tục sử dụng';
    wsMain.getRange(LUANSIM_CELL).setValue(msg);
    SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);

    return;
  }
  */
  
  wsMain.getRange(LUANSIM_CELL).setValue('ĐANG PHÂN TÍCH NHU CẦU...');
  SpreadsheetApp.getActive().toast('ĐANG PHÂN TÍCH NHU CẦU...', "Thầy đang nói", 5);
  SpreadsheetApp.flush(); // Đảm bảo thay đổi được hiển thị ngay lập tức

  // 2. Lấy thông tin cá nhân hóa
  const profession = wsMain.getRange(PROFESSION_CELL).getValue();
  const age = wsMain.getRange(AGE_CELL).getValue();
  const gender = wsMain.getRange(GENDER_CELL).getValue();

  let personalizationInfo = "";
  if (profession) personalizationInfo += `Họ là một ${profession}. `;
  if (gender) personalizationInfo += `Giới tính ${gender}. `;
  if (age) personalizationInfo += `Ở độ tuổi ${age}. `;

  // 3. Đọc nhu cầu của khách hàng từ các hộp kiểm
  const selectedDemands = [];
  let demandCount = 0;
  for (let i = CHECKBOX_START_ROW; i <= CHECKBOX_END_ROW; i++) {
    const checkboxCell = wsMain.getRange('G' + i);
    if (checkboxCell.isChecked()) { // Kiểm tra xem ô kiểm có được chọn không
      const demandValue = wsMain.getRange(i, DEMAND_COLUMN_INDEX).getValue(); // Lấy giá trị từ cột G tương ứng
      if (demandValue && demandValue.toString().trim() !== '') {
        selectedDemands.push(demandValue.toString().trim());
        demandCount++;
        if (demandCount >= 3) break; // Chỉ lấy tối đa 3 lựa chọn
      }
    }
  }

  if (selectedDemands.length === 0) {
    const msg = "Chọn tối thiểu 1, tối đa 3 nhu cầu của khách hàng trong các ô từ F" + CHECKBOX_START_ROW + " đến F" + CHECKBOX_END_ROW + ".";
    wsMain.getRange(LUANSIM_CELL).setValue(msg);
    //Logger.log(msg);
    SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
    return;
  }

  // Chuyển mảng nhu cầu thành chuỗi để đưa vào prompt
  const demandsString = selectedDemands.map((demand, index) => `${index + 1}. ${demand}`).join('\n');

  // 4. Đọc dữ liệu "Điểm nổi bật của sim" từ Q23:Q29
  const highlightDataRange = wsMain.getRange(HIGHLIGHT_CELL_START + ':' + HIGHLIGHT_CELL_END);
  const highlightValues = highlightDataRange.getValues();

  let simHighlights = [];
  for (let i = 0; i < highlightValues.length; i++) {
    const highlight = highlightValues[i][0]; // Lấy giá trị từ cột P
    if (highlight && highlight.toString().trim() !== '') {
      simHighlights.push(highlight.toString().trim());
    }
  }
  const simHighlightsString = simHighlights.map((highlight, index) => `${index + 1}. ${highlight}`).join('\n');

  if (!simHighlightsString) {
    const msg = "Sim này Kết cấu chưa chuẩn. Thầy không khen, cũng không chê, thử với sim của Kim Tâm Cát thầy khen liền";
    SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
    return;
  }

  
  let generatedKhenSim = "";
  try {
    
    const payloadForKhenSimGeneration = {
      action: "GetKhenSimAPI", 
      personalizationInfo: personalizationInfo,
      demandsString: demandsString,
      simHighlightsString: simHighlightsString,
      fileId: fileId
    };
    const optionsForKhenSimGeneration = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payloadForKhenSimGeneration),
      muteHttpExceptions: true // Để bắt lỗi HTTP
    };
    
    const urlAPI = PropertiesService.getScriptProperties().getProperty("urlAPI");
    const response = UrlFetchApp.fetch(urlAPI, optionsForKhenSimGeneration);
    //Logger.log(response);
    const responseCode = response.getResponseCode();
    ketQua = JSON.parse(response.getContentText());
    wsMain.getRange(LUANSIM_CELL).setValue(ketQua.textResult);
    wsMain.getRange(API_LEFT_NUM_CELL).setValue(ketQua.maxUsageAI - ketQua.curentUsageAI);
    SpreadsheetApp.getActive().toast("KHEN SIM THÀNH CÔNG", "Thầy đang nói", 5);
    if (responseCode >= 400) {
      throw new Error(`Lỗi khi tạo luận giải Khen Sim từ File B (${responseCode}): ${ketQua}`);
    }

  } catch (err) {
    const errorMsg = `Lỗi khi tạo luận giải Khen Sim: ${err.message}`;
    wsMain.getRange(LUANSIM_CELL).setValue(errorMsg);
    
    //Browser.msgBox("Lỗi", errorMsg, Browser.Buttons.OK);
  }
}

function BienHoa() {
  

    wsMain.getRange(BIENHOA_CELL).setValue("ĐANG TẢI VUI LÒNG CHỜ...");
    SpreadsheetApp.getActive().toast("ĐANG TẢI VUI LÒNG CHỜ...", "Thầy đang nói", 5);

    wsMain.getRange(LUANSIM_CELL).clearContent();
    const sdt = wsMain.getRange("C6:E6").getDisplayValue();
    const ketQua1 = xuLySoDienThoai(sdt); // P3
    const ketQuaGiuSo0 = xuLyGiuMotSo0(sdt); // Q2
    const ketquaBoPhucVi = boPhucVi(ketQuaGiuSo0); // Q3

    wsMain.getRange("P3").setValue(ketQua1);
    wsMain.getRange("Q2").setValue(ketQuaGiuSo0);
    wsMain.getRange("Q3").setValue(ketquaBoPhucVi);

    // Lấy dãy từ E23 -> L23
  const arr = [];
  for (let i = 0; i < 8; i++) {
    arr.push((wsMain.getRange(23, 5 + i).getValue() + "").toUpperCase().trim());
  }

  
 const payloadForTachTuTruong = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      action: "TachBienHoaAPI",
        arr:  arr,
        sdt: sdt,
        fileId: fileId
        
    })
  };

  
  try {
    const urlAPI = PropertiesService.getScriptProperties().getProperty("urlAPI");
    const response = UrlFetchApp.fetch(urlAPI, payloadForTachTuTruong);
    const ketQua = JSON.parse(response.getContentText());
    
    wsMain.getRange(BIENHOA_CELL).setValue(ketQua.textResult);
    wsMain.getRange(API_LEFT_NUM_CELL).setValue(ketQua.maxUsageAI - ketQua.curentUsageAI);
    SpreadsheetApp.getActive().toast("ĐÃ TÁCH TỪ TRƯỜNG SĐT THÀNH CÔNG", "Thầy đang nói", 5);
    
  } catch (err) {
    wsMain.getRange(BIENHOA_CELL).setValue("LỖI KẾT NỐI API: " + err.message);
  }

  
}
//Sự kiện bấm vào checkbox Lưu khách hàng
function onCheckboxLuuKhachHang(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getA1Notation() == 'J1' && range.getValue() == true) {
    range.clearDataValidations();
    range.setValue("⏳");
    saveCustomer();
    SpreadsheetApp.flush(); // đảm bảo cập nhật hiển thị trước
    Utilities.sleep(1000);
    range.insertCheckboxes();
    range.setValue(false);
  }
}
//Sự kiện bấm vào checkbox Tạo Luận Sim
function onCheckboxLuanSoThuCong(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getA1Notation() == 'B27' && range.getValue() == true) {
    wsMain.getRange(LUANSIM_CELL).clearContent();
    range.clearDataValidations();
    range.setValue("⏳");
    LuanSoThuCong();
    SpreadsheetApp.flush(); // đảm bảo cập nhật hiển thị trước
    Utilities.sleep(1000);
    range.insertCheckboxes();
    range.setValue(false);
    
  }
}
//Sự kiện bấm vào checkbox Luận Sim Thông Minh
function onCheckboxLuanSoThongMinh(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getA1Notation() == 'F27' && range.getValue() == true) {
    wsMain.getRange(LUANSIM_CELL).clearContent();
    range.clearDataValidations();
    range.setValue("⏳");
    LuanSoThongMinh();
    SpreadsheetApp.flush(); // đảm bảo cập nhật hiển thị trước
    Utilities.sleep(1000);
    range.insertCheckboxes();
    range.setValue(false);
  }
}
//Sự kiện bấm vào checkbox Khen Sim
function onCheckboxKhenSim(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getA1Notation() == 'J27' && range.getValue() == true) {
    wsMain.getRange(LUANSIM_CELL).clearContent();
    range.clearDataValidations();
    range.setValue("⏳");
    const SimChuanNL = isSimChuanNL();
    if(!SimChuanNL){
        const msg = "Sim này bị vỡ kết cấu rồi (XẤU), không khen được, cố khen lại bảo thầy ba phải. Bấm nút Luận sim để hiểu thêm con nhé!";
        wsMain.getRange(LUANSIM_CELL).setValue(msg);
        SpreadsheetApp.getActive().toast(msg, "Thầy đang nói", 5);
        
        SpreadsheetApp.flush(); // đảm bảo cập nhật hiển thị trước
        Utilities.sleep(1000);
        range.insertCheckboxes();
        range.setValue(false);
        return;
    }  
    KhenSim();

    SpreadsheetApp.flush(); // đảm bảo cập nhật hiển thị trước
    Utilities.sleep(1000);
    range.insertCheckboxes();
    range.setValue(false);
  }
}

function onEditBienHoa(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() === MAIN_SHEET_NAME) {
    // Nếu sửa bất kỳ ô nào từ E23 → L23 thì chạy
    const row = range.getRow();
    const col = range.getColumn();

    if (row === 6 && col >= 3 && col <= 5) {
      BienHoa(); // gọi hàm chính
    }
  }
}
/*
function TaoTriggerAll() {
CacheService.getScriptCache().remove("urlAPI_1_B");
const url = getMasterSettingWithCache(1, 'B');

  //Trigger
  const triggersHienTai = ScriptApp.getProjectTriggers();
  const ui = SpreadsheetApp.getUi();
  const danhSachTriggerCanTao = [
    "onEditBienHoa",
    "onCheckboxLuanSoThongMinh",
    "onCheckboxLuanSoThuCong",
    "onCheckboxKhenSim",
    "onCheckboxLuuKhachHang"
  ];

  danhSachTriggerCanTao.forEach(tenHam => {
    const daTonTai = triggersHienTai.some(trigger =>
      trigger.getHandlerFunction() === tenHam
    );

    if (!daTonTai) {
      ScriptApp.newTrigger(tenHam)
        .forSpreadsheet(wsMain)
        .onEdit()
        .create();

      //Logger.log(`Đã tạo trigger: ${tenHam}`);
    } else {
      //Logger.log(`Trigger đã tồn tại: ${tenHam}`);
    }
  });
  ui.alert("Đã tạo thành công Trigger chạy tự động luận sim");

}
*/
function isSimChuanNL(){
  const HoaHungvoiKhong = wsMain.getRange(HOAHUNG_CELL).getValues();
  const BienHoa = wsMain.getRange(BIENHOA_CELL).getValue();
  const isHoaHungRong = HoaHungvoiKhong.flat().every(cell => cell === "" || cell === null);
  const isBienHoaRong = BienHoa === "" || BienHoa === null || BienHoa.toLowerCase().trim().includes('đuôi sinh khí');
  if (isHoaHungRong  && isBienHoaRong)return true;
  return false;

}
/*
function onOpen() {
    //Welcome
  //const userEmail = Session.getActiveUser().getEmail();
  //wsMain.getRange("C25").setValue(`Xin chào ${userEmail}`);
  try {
    const url = getMasterSettingWithCache(1, 'B');
    // Đã lưu vào ScriptProperties bên trong hàm rồi
    Logger.log("Lấy URL từ master thành công: " + urlAPI);
  } catch (e) {
    Logger.log("Lỗi khi lấy URL từ Master: " + e.message);
  }
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ShiFu Menu')
      //.addItem('Cho phép tạo luận Biến Hóa', 'BienHoa')
      .addItem('Cho phép tạo Luận sim tự động', 'TaoTriggerAll')
      .addToUi();
  
  
}
*/
function getUrlAPIFromProperties() {
  return PropertiesService.getScriptProperties().getProperty("urlAPI");
}
/** 


function runCopyMainToAnother() {
  const ui = SpreadsheetApp.getUi();
  // Xác nhận lần 1: Cảnh báo chung
  let response = ui.alert(
      'XÁC NHẬN BẮT ĐẦU',
      'Script sắp quét và GHI ĐÈ sheet MAIN trên các file khác, sau đó GỬI EMAIL THÔNG BÁO đến các editor của file đó.\n\nAnh có muốn tiếp tục không?',
      ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    Browser.msgBox('Đã hủy theo yêu cầu của anh.');
    return;
  }
  
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById("1ObkdGAzAhudQxikrR0t1QwJkGesRvqQgtmlO3BWHQsw");

    SpreadsheetApp.flush();
    
    // Tìm kiếm tất cả các file thỏa mãn điều kiện
    const targetFiles = getFilesFromExternalIdList(); // Vẫn sử dụng hàm này

    if (targetFiles.length === 0) {
      Browser.msgBox('Không tìm thấy file', 'Không tìm thấy file Google Sheet nào có ID trong cột C để xử lý.', Browser.Buttons.OK);
      return;
    }

    // Xác nhận lần 2: Hiển thị danh sách file tìm được
    const fileNames = targetFiles.map(f => f.getName()).join('\n');
    response = ui.alert(
        `Đã tìm thấy ${targetFiles.length} file. XÁC NHẬN GHI ĐÈ?`,
        'Các file sau sẽ bị GHI ĐÈ sheet MAIN và các editor sẽ nhận được email:\n\n' + fileNames + '\n\nAnh có chắc chắn muốn tiếp tục không?',
        ui.ButtonSet.YES_NO);

    // Chỉ thực hiện khi người dùng xác nhận lần cuối
    if (response === ui.Button.YES) {
      let successCount = 0;
      for (const file of targetFiles) {
        // Hàm con sẽ thực hiện toàn bộ tác vụ cho từng file
        if (processTargetFile(sourceSpreadsheet, file.getId())) {
          successCount++;
        }
      }
      Browser.msgBox('Hoàn tất', `Đã xử lý ${targetFiles.length} file. Ghi đè và gửi email thành công cho ${successCount} file.`, Browser.Buttons.OK);
    } else {
      Browser.msgBox('Đã hủy.', 'Thao tác hàng loạt đã được hủy.', Browser.Buttons.OK);
    }

  } catch (e) {
    Browser.msgBox('Lỗi!', e.message, Browser.Buttons.OK);
  }
}


function findTargetFilesInSameFolder(sourceSpreadsheet) {
  const sourceFileId = sourceSpreadwsMain.getId();
  const sourceFile = DriveApp.getFileById(sourceFileId);
  const parentFolders = sourceFile.getParents();

  if (!parentFolders.hasNext()) {
    throw new Error('File nguồn không nằm trong thư mục nào. Vui lòng đặt File A vào một thư mục cùng với các file đích.');
  }

  // Lấy thư mục cha đầu tiên của file nguồn
  const parentFolder = parentFolders.next();
  Logger.log(`Đang quét thư mục: "${parentFolder.getName()}"`);

  const emailRegex = /[\w.-]+@[\w.-]+\.\w+/; // Mẫu nhận dạng email
  const filesInFolder = parentFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const targetFiles = [];

  while (filesInFolder.hasNext()) {
    const file = filesInFolder.next();
    // Bỏ qua chính file nguồn
    if (file.getId() === sourceFileId) {
      continue;
    }
    // Kiểm tra tên file có chứa email không
    if (emailRegex.test(file.getName())) {
      targetFiles.push(file);
    }
  }
  return targetFiles;
}

function getFilesFromExternalIdList() {
  // --- CẤU HÌNH ---
  // ID của bảng tính Google Sheet chứa danh sách ID file trong cột C
  const EXTERNAL_SPREADSHEET_ID = "1UvQOYMpF2FqFws4brrHlbH16-QItBpeTNZqV251YYwg"; 
  // Tên của sheet trong bảng tính đó chứa danh sách ID file
  const TARGET_SHEET_NAME = "MAIN"; 
  // Chỉ số cột chứa ID file (Cột C là 3, vì A=1, B=2, C=3)
  const ID_COLUMN_INDEX = 3; 
  // Hàng bắt đầu đọc ID file (bạn yêu cầu từ dòng thứ 2)
  const START_ROW = 2; 
  // --- KẾT THÚC CẤU HÌNH ---

  const targetFiles = []; // Mảng để lưu trữ các đối tượng File tìm được

  try {
    // 1. Mở bảng tính Google Sheet bên ngoài bằng ID
    const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
    
    // 2. Lấy sheet cụ thể từ bảng tính đó
    const targetSheet = externalSpreadwsMain.getSheetByName(TARGET_SHEET_NAME);

    if (!targetSheet) {
      Logger.log(`Lỗi: Không tìm thấy sheet "${TARGET_SHEET_NAME}" trong spreadsheet ID "${EXTERNAL_SPREADSHEET_ID}".`);
      SpreadsheetApp.getUi().alert(`Lỗi: Không tìm thấy sheet "${TARGET_SHEET_NAME}" trong bảng tính nguồn. Vui lòng kiểm tra lại tên wsMain.`);
      return [];
    }

    // 3. Xác định hàng cuối cùng có dữ liệu trong sheet
    const lastRow = targetwsMain.getLastRow();

    // 4. Kiểm tra xem có ID nào để đọc không
    if (lastRow < START_ROW) {
      Logger.log(`Không có ID file nào trong cột C của sheet "${TARGET_SHEET_NAME}" (bắt đầu từ dòng ${START_ROW}).`);
      return [];
    }

    // 5. Lấy tất cả các giá trị ID từ vùng dữ liệu đã chỉ định (cột C từ dòng 2 đến hết)
    const idRange = targetwsMain.getRange(START_ROW, ID_COLUMN_INDEX, lastRow - START_ROW + 1, 1);
    const fileIds = idRange.getValues(); // Trả về một mảng 2D (ví dụ: [["id1"], ["id2"], ...])

    Logger.log(`Tìm thấy ${fileIds.length} ID tiềm năng từ sheet "${TARGET_SHEET_NAME}".`);

    // 6. Duyệt qua từng ID và cố gắng lấy đối tượng File tương ứng
    fileIds.forEach(row => {
      const fileId = String(row[0]).trim(); // Lấy ID từ mảng con và loại bỏ khoảng trắng thừa

      if (fileId) { // Đảm bảo ID không rỗng
        try {
          const file = DriveApp.getFileById(fileId);
          // Tùy chọn: Bạn có thể thêm các kiểm tra khác về file ở đây nếu cần
          // Ví dụ: kiểm tra file.getMimeType() hoặc file.getName()
          targetFiles.push(file); // Thêm đối tượng File vào mảng kết quả
        } catch (fileError) {
          // Log lỗi nếu không thể truy cập file (ví dụ: ID không hợp lệ, không có quyền truy cập, file đã bị xóa)
          Logger.log(`Cảnh báo: Không thể lấy file với ID "${fileId}". Lỗi: ${fileError.message}`);
        }
      }
    });

  } catch (spreadsheetError) {
    // Bắt lỗi nếu không thể mở bảng tính nguồn (ví dụ: ID sai, không có quyền truy cập vào bảng tính đó)
    Logger.log(`Lỗi khi truy cập bảng tính ID "${EXTERNAL_SPREADSHEET_ID}": ${spreadsheetError.message}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi truy cập bảng tính nguồn: ${spreadsheetError.message}. Vui lòng kiểm tra ID và quyền truy cập.`);
    return [];
  }

  Logger.log(`Đã tìm thấy tổng cộng ${targetFiles.length} file hợp lệ.`);
  return targetFiles;
}



function processTargetFile(sourceSpreadsheet, destinationFileId) { // Đã sửa biến
  const SHEET_MAIN_NAME = 'MAIN';
  const SHEET_MOBILE_NAME = 'MAIN MOBILE';
  const TRIGGER_FUNCTION_NAME = 'onEditBienHoa'; // Tên hàm trigger bạn muốn tạo/cập nhật

  try {
    const destinationSpreadsheet = SpreadsheetApp.openById(destinationFileId);
    
    // --- 1. Xóa các sheet cũ "MAIN" và "MAIN MOBILE" trên file đích nếu tồn tại ---
    let oldMainSheet = destinationSpreadwsMain.getSheetByName(SHEET_MAIN_NAME);
    if (oldMainSheet) {
      destinationSpreadwsMain.deleteSheet(oldMainSheet);
      Logger.log(`Đã xóa sheet cũ "${SHEET_MAIN_NAME}" trên file đích ID: ${destinationFileId}`);
    }
    let oldMobileSheet = destinationSpreadwsMain.getSheetByName(SHEET_MOBILE_NAME);
    if (oldMobileSheet) {
      destinationSpreadwsMain.deleteSheet(oldMobileSheet);
      Logger.log(`Đã xóa sheet cũ "${SHEET_MOBILE_NAME}" trên file đích ID: ${destinationFileId}`);
    }

    // --- 2. Copy các sheet "MAIN" và "MAIN MOBILE" mới từ file nguồn sang file đích ---
    // Đã sửa 'sourceSpreadsheets' thành 'sourceSpreadsheet'
    const sourceMainSheet = sourceSpreadwsMain.getSheetByName(SHEET_MAIN_NAME); 
    const sourceMobileSheet = sourceSpreadwsMain.getSheetByName(SHEET_MOBILE_NAME);

    if (!sourceMainSheet) {
      throw new Error(`Không tìm thấy sheet nguồn "${SHEET_MAIN_NAME}" trong file nguồn.`);
    }
    if (!sourceMobileSheet) {
      throw new Error(`Không tìm thấy sheet nguồn "${SHEET_MOBILE_NAME}" trong file nguồn.`);
    }

    // Copy MAIN sheet và đặt tên lại
    const newMainSheet = sourceMainwsMain.copyTo(destinationSpreadsheet);
    newMainwsMain.setName(SHEET_MAIN_NAME);
    // Sao chép các cài đặt bảo vệ từ sheet nguồn sang sheet đích mới
    copyProtections(sourceMainSheet, newMainSheet); 
    Logger.log(`Đã copy và đặt tên lại sheet "${SHEET_MAIN_NAME}" từ nguồn sang đích.`);

    // Copy MAIN MOBILE sheet và đặt tên lại
    const newMobileSheet = sourceMobilewsMain.copyTo(destinationSpreadsheet);
    newMobilewsMain.setName(SHEET_MOBILE_NAME);
    // Sao chép các cài đặt bảo vệ từ sheet nguồn sang sheet đích mới
    copyProtections(sourceMobileSheet, newMobileSheet);
    Logger.log(`Đã copy và đặt tên lại sheet "${SHEET_MOBILE_NAME}" từ nguồn sang đích.`);

    // --- 3. Thay đổi thứ tự Sheet trên file đích ---
    // Di chuyển sheet "MAIN" về vị trí đầu tiên
    destinationSpreadwsMain.setActiveSheet(newMainSheet); // Kích hoạt sheet để có thể di chuyển
    destinationSpreadwsMain.moveActiveSheet(1);
    Logger.log(`Đã di chuyển sheet "${SHEET_MAIN_NAME}" về vị trí đầu tiên.`);

    // Di chuyển sheet "MAIN MOBILE" về vị trí thứ hai
    destinationSpreadwsMain.setActiveSheet(newMobileSheet); // Kích hoạt sheet để có thể di chuyển
    destinationSpreadwsMain.moveActiveSheet(2);
    Logger.log(`Đã di chuyển sheet "${SHEET_MOBILE_NAME}" về vị trí thứ hai.`);

    // --- 4. Cập nhật trigger trên file đích ---
    // Xóa tất cả các trigger cũ có cùng tên hàm cho spreadsheet đích này
    const allTriggers = ScriptApp.getUserTriggers(destinationSpreadsheet);
    for (const trigger of allTriggers) {
      if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`Đã xóa trigger cũ "${TRIGGER_FUNCTION_NAME}" trên file đích ID: ${destinationFileId}`);
      }
    }
    // Tạo mới trigger onEdit cho file đích
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .forSpreadsheet(destinationSpreadsheet)
      .onEdit()
      .create();
    Logger.log(`Đã tạo mới trigger "${TRIGGER_FUNCTION_NAME}" trên file đích ID: ${destinationFileId}`);

    // --- 5. Gửi email thông báo ---
    const editors = destinationSpreadwsMain.getEditors();
    // Lọc bỏ các email rỗng hoặc không hợp lệ trước khi gửi
    const recipientEmails = editors.map(editor => editor.getEmail()).filter(email => email).join(',');
    
    if (recipientEmails) {
      const subject = `Thông báo: File "${destinationSpreadwsMain.getName()}" đã được cập nhật`;
      const body = `Chào bạn,\n\nCông cụ Luận Sim đã có thêm sheet MAIN MOBILE dành cho phiên bản thân thiện hơn khi sử dụng trên điện thoại.\n\n- Tên File: ${destinationSpreadwsMain.getName()}\n- Link truy cập: ${destinationSpreadwsMain.getUrl()}\n\nĐây là email tự động, vui lòng không trả lời.`;
      
      MailApp.sendEmail(recipientEmails, subject, body);
      Logger.log(`Đã gửi email đến: ${recipientEmails}`);
    } else {
      Logger.log('Không tìm thấy địa chỉ email người chỉnh sửa để gửi thông báo.');
    }

    Logger.log(`Xử lý thành công file ID: ${destinationFileId}`);
    return true;

  } catch (e) {
    Logger.log(`Gặp lỗi khi xử lý file ID ${destinationFileId}: ${e.message}`);
    // Hiển thị thông báo lỗi cho người dùng
    SpreadsheetApp.getUi().alert(`Gặp lỗi khi xử lý file ID ${destinationFileId}: ${e.message}`);
    return false;
  }
}

function copyProtectionsToAllRelevantFiles() {
  // Lấy Spreadsheet hiện tại (là File A của bạn)
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSpreadsheetId = sourceSpreadwsMain.getId();
  const sourceSheetName = 'MAIN MOBILE';

  // Lấy sheet nguồn
  const sourceSheet = sourceSpreadwsMain.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('Lỗi', `Không tìm thấy sheet "${sourceSheetName}" trong File nguồn.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Lấy thông tin chủ sở hữu của File A (là người tạo file).
  const owner = sourceSpreadwsMain.getOwner();
  // Lấy thông tin người đang chạy script (là bạn, có thể khác chủ sở hữu nếu file được chia sẻ).
  const me = Session.getEffectiveUser();

  // Lấy tất cả các cài đặt bảo vệ trên dải ô và toàn trang từ sheet nguồn
  const sourceRangeProtections = sourcewsMain.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const sourceSheetProtections = sourcewsMain.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  let filesProcessedCount = 0;
  let sheetsUpdatedCount = 0;

  try {
    // Lấy tất cả các thư mục chứa File A
    const parentFolders = DriveApp.getFileById(sourceSpreadsheetId).getParents();
    let folderId = null;

    // Lấy ID của thư mục cha đầu tiên (giả định File A nằm trong ít nhất 1 thư mục)
    if (parentFolders.hasNext()) {
      folderId = parentFolders.next().getId();
    } else {
      SpreadsheetApp.getUi().alert('Lỗi', 'File nguồn không nằm trong thư mục nào. Không thể xác định thư mục để tìm các file khác.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const folder = DriveApp.getFolderById(folderId);
    const filesInFolder = folder.getFilesByType(MimeType.GOOGLE_SHEETS); // Lấy tất cả Google Sheets trong thư mục

    while (filesInFolder.hasNext()) {
      const file = filesInFolder.next();

      // Bỏ qua File nguồn để không xử lý lại chính nó
      if (file.getId() === sourceSpreadsheetId) {
        continue;
      }

      let destinationSpreadsheet;
      try {
        destinationSpreadsheet = SpreadsheetApp.openById(file.getId());
      } catch (e) {
        Logger.log(`Không thể mở file đích ${file.getName()} (ID: ${file.getId()}): ${e.message}`);
        continue; // Bỏ qua file nếu không thể mở (ví dụ: không có quyền)
      }

      const destinationSheet = destinationSpreadwsMain.getSheetByName(sourceSheetName);

      // Chỉ xử lý nếu sheet đích tồn tại
      if (destinationSheet) {
        filesProcessedCount++;
        sheetsUpdatedCount++;
        Logger.log(`Đang xử lý sheet "${sourceSheetName}" trong file: ${destinationSpreadwsMain.getName()} (ID: ${destinationSpreadwsMain.getId()})`);

        // --- Xóa tất cả các bảo vệ hiện có trên sheet đích trước khi sao chép
        // để tránh chồng chong hoặc lỗi nếu đã có bảo vệ cũ.
        const currentSheetProtections = destinationwsMain.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        for(const prot of currentSheetProtections) {
          try {
            prot.remove();
          } catch (e) {
            Logger.log(`Không thể xóa bảo vệ toàn sheet của file ${destinationSpreadwsMain.getName()}: ${e.message}`);
          }
        }
        const currentRangeProtections = destinationwsMain.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for(const prot of currentRangeProtections) {
          try {
            prot.remove();
          } catch (e) {
            Logger.log(`Không thể xóa bảo vệ dải ô của file ${destinationSpreadwsMain.getName()}: ${e.message}`);
          }
        }

        // --- SAO CHÉP BẢO VỆ TRÊN TỪNG DẢI Ô (RANGE PROTECTIONS) ---
        for (const protection of sourceRangeProtections) {
          try {
            const range = destinationwsMain.getRange(protection.getRange().getA1Notation());
            const newProtection = range.protect();

            newProtection.setDescription(protection.getDescription());
            newProtection.setWarningOnly(protection.isWarningOnly());
            newProtection.removeEditors(newProtection.getEditors());
            newProtection.addEditors([owner, me]);
          } catch (e) {
            Logger.log(`Lỗi khi sao chép bảo vệ dải ô ${protection.getRange().getA1Notation()} vào file ${destinationSpreadwsMain.getName()}: ${e.message}`);
          }
        }

        // --- SAO CHÉP BẢO VỆ TOÀN TRANG TÍNH (SHEET PROTECTION) ---
        if (sourceSheetProtections.length > 0) {
          try {
            const protection = sourceSheetProtections[0]; // Giả định chỉ có 1 bảo vệ trên mỗi sheet
            const newProtection = destinationwsMain.protect();

            newProtection.setDescription(protection.getDescription());
            newProtection.setWarningOnly(protection.isWarningOnly());
            newProtection.removeEditors(newProtection.getEditors());
            newProtection.addEditors([owner, me]);

            // Sao chép các dải ô không được bảo vệ (Unprotected Ranges)
            const unprotectedRanges = protection.getUnprotectedRanges();
            if (unprotectedRanges.length > 0) {
              const newUnprotectedRanges = unprotectedRanges.map(r => destinationwsMain.getRange(r.getA1Notation()));
              newProtection.setUnprotectedRanges(newUnprotectedRanges);
            }
          } catch (e) {
            Logger.log(`Lỗi khi sao chép bảo vệ toàn trang vào file ${destinationSpreadwsMain.getName()}: ${e.message}`);
          }
        }
      }
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Lỗi chung', `Có lỗi xảy ra khi duyệt thư mục hoặc xử lý file: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  Logger.log(`Đã xử lý ${filesProcessedCount} file và cập nhật bảo vệ cho ${sheetsUpdatedCount} sheet "${sourceSheetName}".`);
}


function copyProtections(sourceSheet, destinationSheet) {
  // Lấy thông tin chủ sở hữu của file hiện tại (File mà script đang chạy - File A).
  // Đã được định nghĩa ở ngoài hàm này để tránh gọi lại nhiều lần và đảm bảo chính xác chủ sở hữu của File A.
  const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  // Lấy thông tin người đang chạy script.
  const me = Session.getEffectiveUser();

  // --- 1. SAO CHÉP BẢO VỆ TRÊN TỪNG DẢI Ô (RANGE PROTECTIONS) ---
  const rangeProtections = sourcewsMain.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const protection of rangeProtections) {
    const range = destinationwsMain.getRange(protection.getRange().getA1Notation());
    const newProtection = range.protect(); // Tạo một bảo vệ mới

    // Sao chép các thuộc tính cơ bản
    newProtection.setDescription(protection.getDescription());
    newProtection.setWarningOnly(protection.isWarningOnly());
    
    // Xóa tất cả người chỉnh sửa mặc định (nếu có)
    newProtection.removeEditors(newProtection.getEditors());
    
    // Thêm CHỈ chủ sở hữu file và người chạy script vào danh sách được phép chỉnh sửa
    newProtection.addEditors([owner, me]);
  }

  // --- 2. SAO CHÉP BẢO VỆ TOÀN TRANG TÍNH (SHEET PROTECTION) ---
  const sheetProtections = sourcewsMain.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (sheetProtections.length > 0) {
    const protection = sheetProtections[0]; // Giả định chỉ có 1 bảo vệ trên mỗi sheet
    const newProtection = destinationwsMain.protect();

    // Sao chép các thuộc tính cơ bản
    newProtection.setDescription(protection.getDescription());
    newProtection.setWarningOnly(protection.isWarningOnly());
    
    // Xóa tất cả người chỉnh sửa mặc định
    newProtection.removeEditors(newProtection.getEditors());

    // Thêm CHỈ chủ sở hữu file và người chạy script vào danh sách được phép chỉnh sửa
    newProtection.addEditors([owner, me]);
    
    // Sao chép các dải ô không được bảo vệ (Unprotected Ranges)
    const unprotectedRanges = protection.getUnprotectedRanges();
    if (unprotectedRanges.length > 0) {
      const newUnprotectedRanges = unprotectedRanges.map(r => destinationwsMain.getRange(r.getA1Notation()));
      newProtection.setUnprotectedRanges(newUnprotectedRanges);
    }
  }
}

*/
function setTest(fileID,sString) {
  const sheet = SpreadsheetApp.openById("1UvQOYMpF2FqFws4brrHlbH16-QItBpeTNZqV251YYwg").getSheetByName('MAIN');
  const data = wsMain.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === fileID) { // Cột C (index 2)
      const currentNum = parseInt(data[i][5], 10) || 0; // Cột F (index 5)
      //const today = new Date(); // ngày giờ hiện tại

      wsMain.getRange(i + 1, 6).setValue(currentNum + 1); // cập nhật cột F
      wsMain.getRange(i + 1, 7).setValue(sString); // cập nhật cột G

      return;
    }
  }

  throw new Error("setTest Không tìm thấy fileID trong bảng MASTER."+fileID );
}
