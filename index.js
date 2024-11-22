

function showLoading() {
    document.getElementById('loadingIcon').style.display = 'block';
}

function hideLoading() {
    document.getElementById('loadingIcon').style.display = 'none';
}

function showLoadingSearch() {
    document.getElementById('loadingIconSearch').style.display = 'block';
}
function hideLoadingSearch() {
    document.getElementById('loadingIconSearch').style.display = 'none';
}

const apiKey = 'AIzaSyCUkklmX-ewNdmcDfJCb5FImBwfN0F4wjg';
const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
const rangesheet1 = 'data!A3:ZZ';
const rangesheet2 = 'program!A2:ZZ';
const rangesheet3 = 'register!A2:ZZ';
const rangesheet4 = 'register!A2:A';
const rangesheet5 = 'sticker!A2:ZZ';
const rangesheet6 = 'specimencount!A2:ZZ';
const rangesheet7 = 'specimencount!A:A';
const rangesheet8 = 'program!A2:ZZ';
const rangesheet9 = 'sticker!A2:ZZ';
const rangesheet10 = 'register!A2:ZZ';
const rangesheet11 = 'register!K2:KK';
const rangesheet12 = 'specimencount!B1:EE';





async function addRegistrationData() {
    try {
        // เรียกฟังก์ชันแสดงลำดับถัดไป
        displayNextNumber();

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`;
        
        // ดึงข้อมูลทั้งหมดจาก Google Sheets
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error('Network response was not ok: ' + response.statusText);
        }

        const sheetData = await response.json();

        // ตรวจสอบว่ามี idcard ซ้ำหรือไม่
        const rows = sheetData.values || [];
        const isDuplicate = rows.some(row => row[3] === idcard); // สมมุติว่า idcard อยู่ในคอลัมน์ที่ 4 (index 3)

        if (isDuplicate) {
            // แจ้งเตือนว่ามี idcard นี้ลงทะเบียนแล้ว
            Swal.fire({
                icon: 'error',
                title: 'Oops...',
                text: 'ID นี้ลงทะเบียนแล้ว!'
            });
            return;
        }

    var number = document.getElementById('numb').textContent.trim();
    var regisid = document.getElementById('registernumber').textContent.trim();
    var name = document.getElementById('name').textContent.trim();
    var idcard = document.getElementById('idcard').textContent.trim();
    var sexage = document.getElementById('age').textContent.trim();
    var card = document.getElementById('card').textContent.trim();
    var depart = document.getElementById('depart').textContent.trim();
    var birth = document.getElementById('birthday').textContent.trim();
    var prog = document.getElementById('program').textContent.trim();
    var date = document.getElementById('datetime').textContent.trim();
    var desc = document.getElementById('desc').textContent.trim();


        const rowData = [[number, regisid, name, idcard, card, depart, sexage, birth, prog, date, desc]];
        console.log(rowData);
        // URL ของ Google Apps Script Web App ที่ Deploy ไว้
        const url1 = 'https://script.google.com/macros/s/AKfycbyXUGV1bM84mVLRy2DZNLIz0uSf5N2xgG_cDQ4nNMAqo7oVh_GJSz6yS1HkYAnAfLHW2Q/exec';
        console.log(url1);
        // เตรียมข้อมูลที่ต้องการส่งไปยัง Google Apps Script
        const postData = {
            action: 'addRegistration',
            rowData: rowData
        };

        // ส่งข้อมูลไปยัง Google Apps Script
        const postResponse = await fetch(url1, {
            method: 'POST',
            body: JSON.stringify(postData)
        });

        const result = await postResponse.json();
        if (result.success) {
            Swal.fire({
                position: "center",
                icon: "success",
                title: "เพิ่มข้อมูลสำเร็จ",
                showConfirmButton: false,
                timer: 1500
            });

            // เรียกฟังก์ชันภายในเพิ่มเติม
            await addRegistrationDataInner();
            loadAllRecords();
        } else {
            throw new Error(result.error || 'Failed to add data');
        }

    } catch (error) {
        console.error('Error:', error);
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'เกิดข้อผิดพลาดในการลงทะเบียน! กรุณาลองอีกครั้ง.'
        });
    }
}

async function addRegistrationDataInner() {
    try {
        const numr = document.getElementById('datetime').textContent.trim();
        const regisid = document.getElementById('registernumber').textContent.trim();
        const name = document.getElementById('name').textContent.trim();
        const type = "1ลงทะเบียน";
        const spec = 10;

        // สร้างข้อมูลที่ต้องการเพิ่ม
        const rowData = [[numr, regisid, name, spec, type]];

        // URL ของ Google Apps Script Web App ที่ Deploy ไว้
        const url2 = 'https://script.google.com/macros/s/AKfycbyXUGV1bM84mVLRy2DZNLIz0uSf5N2xgG_cDQ4nNMAqo7oVh_GJSz6yS1HkYAnAfLHW2Q/exec';

        // เตรียมข้อมูลที่ต้องการส่งไปยัง Google Apps Script
        const postData = {
            action: 'addRegistrationInner',
            rowData: rowData
        };

        // ส่งข้อมูลไปยัง Google Apps Script
        const response = await fetch(url2, {
            method: 'POST',
            body: JSON.stringify(postData)
        });

        const result = await response.json();

        if (result.success) {
            Swal.fire({
                position: "center",
                icon: "success",
                title: "เพิ่มข้อมูลสำเร็จ",
                showConfirmButton: false,
                timer: 1500
            });
        } else {
            throw new Error(result.error || 'Failed to add data');
        }
    } catch (error) {
        console.error('Error in adding data to Google Sheets:', error);
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'เกิดข้อผิดพลาดในการลงทะเบียน!'
        });
    }
}


async function buildSticker() {
    try {
        // รับค่า input จากผู้ใช้งาน
        const program = document.getElementById('newprogram').value.trim();
        const regisid = document.getElementById('newid').value.trim();
        const name = document.getElementById('newname').value.trim();
        const idcard = document.getElementById('newidcard').value.trim();

        // ตรวจสอบข้อมูลที่จำเป็น
        if (!program || !regisid || !name || !idcard) {
            Swal.fire({
                icon: 'warning',
                title: 'ข้อมูลไม่ครบถ้วน',
                text: 'กรุณากรอกข้อมูลให้ครบถ้วน',
            });
            return;
        }

        // URL สำหรับเรียก Google Sheets API
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet8}?key=${apiKey}`;

        // ดึงข้อมูลจาก Google Sheets
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`Network response was not ok: ${response.statusText}`);
        }

        const sheetData = await response.json();
        const rows = sheetData.values || [];

        // ค้นหาแถวที่ตรงกับโปรแกรมที่เลือก
        const matchingRows = rows.filter(row => row[0] === program);

        if (matchingRows.length === 0) {
            Swal.fire({
                icon: 'error',
                title: 'ไม่พบข้อมูล',
                text: 'ไม่พบโปรแกรมที่เลือกในฐานข้อมูล',
            });
            return;
        }

       
        const rowData = matchingRows.map(row => {
            const method = row[2] || 'Unknown method';
            const methodid = row[3] || 'Unknown methodid';
            const custom = row[4] || 'Unknown custom';

            const barcodesticker = `*${regisid}${method}*`;
            const stickerid = `${regisid}${program}`;

            return [regisid, barcodesticker, stickerid, name, custom, method];
        });

        console.log('Prepared data for stickers:', rowData);

        // เตรียมข้อมูลสำหรับส่งไปยัง Google Apps Script
        const scriptURL = 'https://script.google.com/macros/s/AKfycbwmsn50jXDf6ygVN6MkDf1Q139m-bj17w7AntkKt-x8YOnh46aGa9RIb6zDONLJp3ZhSQ/exec';
        const postData = {
            action: 'buildSticker',
            rowData: rowData,
        };

        // ส่งข้อมูลไปยัง Google Apps Script
        const postResponse = await fetch(scriptURL, {
            method: 'POST',
            body: JSON.stringify(postData)
         
        });

        const result = await postResponse.json();

        if (result.success) {
            Swal.fire({
                position: 'center',
                icon: 'success',
                title: 'เพิ่มข้อมูลสำเร็จ',
                showConfirmButton: false,
                timer: 1500,
            });
      
           
           
        } else {
            throw new Error(result.error || 'Failed to save data');
        }
    } catch (error) {
        console.error('Error in buildSticker:', error);
        Swal.fire({
            icon: 'error',
            title: 'เกิดข้อผิดพลาด',
            text: 'ไม่สามารถบันทึกข้อมูลได้ กรุณาลองใหม่อีกครั้ง.',
        });
    }
     closeNewRegister();
     searchData();
}





async function addNewData() {
    try {
        // รับค่าจากฟอร์ม
        const newid = document.getElementById('newid').value.trim();
        const newname = document.getElementById('newname').value.trim();
        const newidcard = document.getElementById('newidcard').value.trim();
        const birthdate = document.getElementById('birthdate').value.trim();
        const newcard = document.getElementById('newcard').value.trim();
        const newdepart = document.getElementById('newdepart').value.trim();
        const newage = document.getElementById('newage').textContent.trim();
        const newprogram = document.getElementById('newprogram').value.trim();

        // ตรวจสอบข้อมูลว่าครบถ้วนหรือไม่
        if (!newid || !newname || !newidcard || !birthdate || !newcard || !newdepart || !newage || !newprogram) {
            Swal.fire({
                icon: 'warning',
                title: 'ข้อมูลไม่ครบถ้วน',
                text: 'กรุณากรอกข้อมูลให้ครบถ้วน'
            });
            return;
        }

        // เตรียมข้อมูลที่ต้องการเพิ่ม
        const rowData = [[newid, newname, newidcard, newcard, newdepart, newage, birthdate, newprogram]];
        console.log(rowData);

        // URL ของ Google Apps Script Web App
        const scriptURL2 = 'https://script.google.com/macros/s/AKfycbyXUGV1bM84mVLRy2DZNLIz0uSf5N2xgG_cDQ4nNMAqo7oVh_GJSz6yS1HkYAnAfLHW2Q/exec';

        
        const postData = {
            action: 'addNewData',
            rowData: rowData
        };

        // ส่งข้อมูลไปยัง Google Apps Script
        const response = await fetch(scriptURL2, {
            method: 'POST',
            body: JSON.stringify(postData)
                   });

        // ตรวจสอบสถานะการตอบกลับ
        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }

        const result = await response.json();

        // ตรวจสอบผลลัพธ์ที่ได้รับ
        if (result.success) {
            Swal.fire({
                position: "center",
                icon: "success",
                title: "เพิ่มข้อมูลสำเร็จ",
                showConfirmButton: false,
                timer: 1500
            });

           const a = document.getElementById('idcard');
           const b = document.getElementById('newidcard').value;
           const c = document.getElementById('desc'); 
           a.innerText = b;
           c.innerText = "เพิ่มชื่อ";
            await buildSticker();
        } else {
            throw new Error(result.error || 'Failed to add data');
        }
    } catch (error) {
        console.error('Error in adding data to Google Sheets:', error);
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'เกิดข้อผิดพลาดในการเพิ่มข้อมูล!'
        });
    }
}


async function addRegistData() {
    try {
        // รับค่าจาก input และ element ต่าง ๆ
        var barinput = document.getElementById('inputbar').value.trim();
        var barcodenewid = document.getElementById('barregisterid').textContent.trim();
        var barcodename = document.getElementById('barname').textContent.trim();
        var inputdate = document.getElementById('datetime').textContent;
        var inputmethodbar = barinput.slice(-2); // ดึง 2 ตัวท้ายของบาร์โค้ด 
        console.log("Method:", inputmethodbar);

        // กำหนดค่า Specimen ตาม method
        var specimen;
        switch (inputmethodbar) {
            case "11":
                specimen = "PE";
                break;
            case "12":
                specimen = "EDTA";
                break;
            case "13":
                specimen = "Urine";
                break;
            case "14":
                specimen = "X Ray";
                break;
            case "15":
                specimen = "EKG";
                break;
            case "20":
                specimen = "naf";
                break;
            case "21":
                specimen = "Clot";
                break;
            case "16":
                specimen = "Audiogram";
                break;
            case "17":
                specimen = "Lung";
                break;
            case "18":
                specimen = "Eyes";
                break;
            case "19":
                specimen = "Muscle";
                break;
            default:
                specimen = "ไม่พบข้อมูล";
                break;
        }

        // เตรียมข้อมูลที่จะส่งไปยัง Google Apps Script
        var data = {
            values: [[inputdate, barcodenewid, barcodename, inputmethodbar, specimen]]
        };
        console.log("Data to be sent:", data);

        // URL ของ Google Apps Script
        const url4 = `https://script.google.com/macros/s/AKfycbyXUGV1bM84mVLRy2DZNLIz0uSf5N2xgG_cDQ4nNMAqo7oVh_GJSz6yS1HkYAnAfLHW2Q/exec`;

        // เตรียมข้อมูลสำหรับ POST request
        const postData = {
            action: 'addRegistData',
            rowData: data
        };

        // ส่งข้อมูลด้วย fetch
        const response = await fetch(url4, {
            method: 'POST',
            body: JSON.stringify(postData)
        });

        // ตรวจสอบว่า response สำเร็จหรือไม่
        if (!response.ok) {
            throw new Error('Network response was not ok: ' + response.statusText);
        }

        // อ่านผลลัพธ์จาก response
        const result = await response.json();
        console.log("Success:", result);

       loadDataTable();

    } catch (error) {
        // เมื่อเกิดข้อผิดพลาดในการบันทึกข้อมูล
        console.error('Error:', error);
        Swal.fire({
            icon: 'error',
            title: 'เกิดข้อผิดพลาด',
            text: 'ไม่สามารถเพิ่มข้อมูลได้'
        });
    }
}



function searchData() {
    const searchKeyElement = document.getElementById('searchKey');

    if (searchKeyElement && searchKeyElement.value === '') {
        searchDataFromId();
    } else {
        searchDataKey();
    }
}

async function searchDataFromId() {
    const searchId = document.getElementById('idcard').textContent.trim(); 

    const url = `https://script.google.com/macros/s/AKfycbyvjMitlmt9ywP_nRWzh3dQyDL-pxQ6o1VbC6-fxpCdn386x8BuFMK5G5Ni9Ncm4zmpvw/exec`; 

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
           body: JSON.stringify({ action: 'searchData', searchId })
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();

        // อ้างอิง element UI
        const regisid = document.getElementById("registernumber");
        const name = document.getElementById("name");
        const card = document.getElementById("card");
        const depart = document.getElementById("depart");
        const age = document.getElementById("age");
        const birthday = document.getElementById("birthday");
        const program = document.getElementById("program");
        const desc = document.getElementById("desc");
        // ล้างผลลัพธ์ก่อนหน้า
        regisid.textContent = "";
        name.textContent = "";
        age.textContent = "";
        birthday.textContent = "";
        program.textContent = "";
        card.textContent = "";
        depart.textContent = "";
         desc.textContent ="";
        if (data.success) {
            // แสดงผลลัพธ์
            const row = data.result;

            regisid.textContent = row[0];
            name.textContent = row[1];
            card.textContent = row[3];
            depart.textContent = row[4];
            age.textContent = row[5];
            birthday.textContent = birthdayFormat(row[6]);
            program.textContent = row[7];
           
            // เรียกฟังก์ชันเพิ่มเติม
            setTimeout(() => {
                searchProgram();
            }, 5000);
            searchPrint();
        } else {
            alert("ไม่พบข้อมูลในระบบ");
        }
    } catch (error) {
        console.error('Error:', error);
        alert("เกิดข้อผิดพลาดในการค้นหาข้อมูล");
    }
}





async function searchDataKey() {
    displayNextNumber();
    const searchKey = document.getElementById('searchKey').value.trim(); // ดึงค่าจาก input และลบช่องว่าง

    const url = `https://script.google.com/macros/s/AKfycbzIt68Ot6cAniCQ610NQ0779fqWnAngLOZYWf2DmQTf-kld2kKh09zsBR7SXiCUuFBh/exec`; // URL ของ Apps Script

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action: 'searchDataKey', searchKey })
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();

        // อ้างอิง element UI
        const regisid = document.getElementById("registernumber");
        const name = document.getElementById("name");
        const idcard = document.getElementById("idcard");
        const card = document.getElementById("card");
        const depart = document.getElementById("depart");
        const age = document.getElementById("age");
        const birthday = document.getElementById("birthday");
        const program = document.getElementById("program");

        // ล้างผลลัพธ์ก่อนหน้า
        regisid.textContent = "";
        name.textContent = "";
        idcard.textContent = "";
        age.textContent = "";
        birthday.textContent = "";
        program.textContent = "";
        card.textContent = "";
        depart.textContent = "";

        if (data.success) {
            // แสดงผลลัพธ์
            const row = data.result;

            regisid.textContent = row[0];
            name.textContent = row[1];
            idcard.textContent = row[2];
            card.textContent = row[3];
            depart.textContent = row[4];
            age.textContent = row[5];
            birthday.textContent = birthdayFormat(row[6]);
            program.textContent = row[7];

            // เรียกฟังก์ชันเพิ่มเติม
            setTimeout(() => {
                searchProgram();
            }, 5000);
            searchPrint();
        } else {
            alert("ไม่พบข้อมูลในระบบ");
        }
    } catch (error) {
        console.error('Error:', error);
        alert("เกิดข้อผิดพลาดในการค้นหาข้อมูล");
    }
}






async function loadAllRecords() {
    const url = `https://script.google.com/macros/s/AKfycbyo_bf9YyMBF7k-5PejJbYC4drUKr5AjeTmuGFg5DUTS2E7mtYQLkIZoB9w0ShY5KzMoA/exec`; // URL ของ Apps Script
    const resultDiv1 = document.getElementById('registeresult');
    resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action: 'loadAllRecords' })
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();

        if (data.success) {
            if (data.values && data.values.length > 0) {
                data.values.forEach(row => {
                    var datetime = document.getElementById('datetime')
                    resultDiv1.innerHTML += `
<tr>
 <th scope="col" class="col-1 text-center">${row[0]}</th>
 <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[1]}</td>
 <td scope="col" class="col-4" style="font-family: sarabun;">${row[2]}</td>
 <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[8]}</td>
 <td scope="col" class="col-3 text-center" style="font-family: sarabun;">${formatDateTime(row[9])}</td>
</tr>
`;
                });
            } else {
                resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">ไม่พบข้อมูล</td></tr>`;
            }
        } else {
            resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">${data.message}</td></tr>`;
        }
    } catch (error) {
        console.error('Error loading all records:', error);
        alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
    }
}




async function searchPrint() {
    const url = `https://script.google.com/macros/s/AKfycbxjGhKa8ELJCe1Bb4QJ6thawbNxrRGTHUpL4CRW7L4bJgRakrEvyfwNxH0S76z_Z5K8bQ/exec`; // URL ของ Apps Script
    const sticker = document.getElementById('sticker');
    const number = document.getElementById('numb').textContent.trim(); // ลบช่องว่าง
    const searchKey = document.getElementById('registernumber').textContent.trim();
    sticker.innerHTML = ''; // ล้างเนื้อหาก่อนแสดงใหม่

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({
                action: 'searchPrint',
                searchKey: searchKey
            }),
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();
        let found = false; // ตรวจสอบว่าพบข้อมูลหรือไม่

        if (data.success && data.values.length > 0) {
            data.values.forEach((row, index) => {
                if (String(row[0]) === String(searchKey)) {
                    found = true; // พบข้อมูล
                    const uniqueId = `barcode-${index}`; // ID ที่ไม่ซ้ำกัน

                    sticker.innerHTML += `
<div class="sticker mb-2">
    <div class="sticker-id mx-2 my-2 text-center"><span>${row[0]}</span></div>
    <div class="sticker-right">
        <div class="barcode-container mt-1">
            <div class="mt-2 ms-2"><span class="libre-barcode-39-extended-regular">${row[1]}</span></div>
        </div>
        <div class="sticker-inner">
            <div class="sticker-id fw-bold" style="font-size: 0.75rem;"><span>${number}</span></div>
            <div class="sticker-right">
                <div class="sticker-name ms-2">${row[3]}</div>
                <div class="sticker-name">
                    <div class="sticker-method ms-2">${row[5]}</div>
                    <div class="sticker-date mt-1">${updateDate()}</div>
                </div>
                <div class="sticker-name ms-4" style="font-size: 0.5rem">${row[4]}</div>
            </div>
        </div>
    </div>
</div>
`;
                }
            });

            if (!found) {
                alert("ไม่พบข้อมูลที่ค้นหา");
            }
        } else {
            alert("ไม่พบข้อมูลที่ค้นหา");
        }
    } catch (error) {
        console.error('Error fetching data:', error);
    }
}

async function loadAllCount() {
    const url = `https://script.google.com/macros/s/AKfycbwmG0sP6B0HBL_y_yI0BP-TSPL77b-a5YhQ5IIk4MyjAdFBv4nH34WmF-TsZPVRb7bOXQ/exec`; // URL ของ Apps Script

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({
                action: 'loadAllCount',
            }),
        });

        if (!response.ok) {
            throw new Error("Network response was not ok");
        }

        const data = await response.json();

        if (data.success) {
            const counts = data.counts;

            document.getElementById('PE').textContent = counts.a || 0;
            document.getElementById('EDTA').textContent = counts.b || 0;
            document.getElementById('urine').textContent = counts.c || 0;
            document.getElementById('xray').textContent = counts.d || 0;
            document.getElementById('ekg').textContent = counts.e || 0;
            document.getElementById('ear').textContent = counts.f || 0;
            document.getElementById('lung').textContent = counts.g || 0;
            document.getElementById('eye').textContent = counts.h || 0;
            document.getElementById('muscle').textContent = counts.i || 0;
            document.getElementById('naf').textContent = counts.j || 0;
            document.getElementById('clot').textContent = counts.k || 0;
            document.getElementById('register').textContent = counts.l || 0;

            clearSpecimen();
            loadDataTable();
        } else {
            console.error('Failed to load counts:', data.message);
        }
    } catch (error) {
        console.error('Error fetching counts:', error);
    }
}


async function loadDataTable() {
    const url = `https://script.google.com/macros/s/AKfycbzHTHNmxWOXrgnvIFrEjC05J0BBh8QQOtOb-nKiAqwM4joaqw_biTqU_Ew2t0UkY2dp1A/exec`; // URL ของ Apps Script

    try {
        // ส่งคำขอไปยัง Apps Script
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({
                action: 'loadDataTable',
            }),
        });

        if (!response.ok) {
            throw new Error("Network response was not ok: " + response.statusText);
        }

        const data = await response.json();

        const resultDiv1 = document.getElementById('specimenresult');
        resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

        if (data.success && data.records.length > 0) {
            const sortedData = data.records.sort((a, b) => {
                return parseInt(b[0], 10) - parseInt(a[0], 10); // เรียงลำดับข้อมูล
            });

            sortedData.forEach(row => {
                if (row[4] !== "1ลงทะเบียน") {
                    resultDiv1.innerHTML += `
                        <tr>
                            <th scope="row" class="col-3 text-center">${row[0]}</th>
                            <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[1] || 'N/A'}</td>
                            <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2] || 'N/A'}</td>
                            <td scope="col" class="text-center" style="font-family: sarabun;">${row[4] || 'N/A'}</td>
                        </tr>`;
                }
            });
        } else {
            resultDiv1.innerHTML = "<tr><td colspan='8' class='text-center'>ไม่พบข้อมูล</td></tr>";
        }

        loadAllCount(); // เรียกฟังก์ชัน loadAllCount หลังจากโหลดข้อมูลเสร็จ
    } catch (error) {
        console.error("Error fetching data:", error);
        alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
    }
}



function sendBarcode() {
    const barcode = document.getElementById('inputbar').value.trim(); // ดึงค่าจาก input และลบช่องว่าง
    const scriptURL = "https://script.google.com/macros/s/AKfycbzhaijq29CYZzzoq9haj3hzwZc6Qwv1Eupu_4Oa7C0zZscb3E0ABg20ZlJVgIn1nL4MPg/exec"; // URL ของ Apps Script Web App

    // ส่งคำขอ POST ไปยัง Apps Script
    fetch(scriptURL, {
        method: 'POST',
        body: JSON.stringify({ barcode })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error("Network response was not ok");
        }
        return response.json();
    })
    .then(data => {
        const baridElement = document.getElementById('barregisterid');
        const barnameElement = document.getElementById('barname');

        // ล้างค่าที่แสดงอยู่ก่อนหน้า
        baridElement.textContent = '';
        barnameElement.textContent = '';

        if (data.status === 'success') {
            // ถ้าพบข้อมูล
            const record = data.record;
            baridElement.textContent = record.barId;
            barnameElement.textContent = record.barName;

            // รอ 1 วินาทีเพื่อเรียก `addRegistData`
            setTimeout(() => {
                addRegistData();
            }, 1000);
        } else {
            // ถ้าไม่พบข้อมูล แสดงแจ้งเตือน
            alert(data.message);
        }
    })
    .catch(error => {
        console.error('Error in fetching barcode data:', error);
        alert('เกิดข้อผิดพลาดในการค้นหาข้อมูล');
    });
}



async function loadAllRecordsSearch() {
    const scriptURL = "https://script.google.com/macros/s/AKfycbzyfS4ZEBGYy7En0Im4HQ6_k46z_0W7fUf6cYLPtaHyNlri-D27PP-yVj_TmUVsnO4hVQ/exec"; // URL ของ Apps Script Web App

    try {
        const response = await fetch(scriptURL); // รอการตอบกลับจาก Web App
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        
        const data = await response.json(); // แปลงข้อมูลที่ได้รับเป็น JSON

        const resultDiv1 = document.getElementById('searchdataload');
        resultDiv1.innerHTML = ''; // ล้างข้อมูลก่อนหน้า

        if (data && data.length > 0) { // ตรวจสอบว่ามีข้อมูลหรือไม่
            data.forEach(row => {
                resultDiv1.innerHTML += `
<tr onclick="copyId(this)" style="cursor: pointer;">
    <th class="column-5 text-center">${row[0]}</th>
    <td class="column-6">${row[1]}</td>
    <td class="column-5 text-center">${row[2]}</td>
    <td class="column-5 text-center">${row[4]}</td>
</tr>`;
            });
            hideLoadingSearch();
        } else {
            resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">ไม่พบข้อมูล</td></tr>`;
        }
    } catch (error) {
        console.error('Error loading all records:', error);
        alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
    }
}



async function searchDataSearch() {
    const searchId = document.getElementById('datasearchid').value.trim(); // ดึงค่าจาก input และลบช่องว่าง

    const scriptURL = "https://script.google.com/macros/s/AKfycbybx_krtf3wBb23RYZifPKT9yFLm-lBjipS5fJ11nvPJqw2Ua2FyrVQK6p0nKUNb0dZ/exec"; // URL ของ Apps Script Web App

    try {
        const response = await fetch(scriptURL); // ใช้ fetch เพื่อดึงข้อมูลจาก Web App
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json(); // แปลงข้อมูลที่ได้รับเป็น JSON

        const resultDiv1 = document.getElementById('searchdataload');
        resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

        if (data && data.length > 0) {
            data.forEach(row => {
                if (row[2] === searchId) {
                    resultDiv1.innerHTML += `
<tr onclick="copyId(this)" style="cursor: pointer;">
    <th class="column-5 text-center">${row[0]}</th>
    <td class="column-6">${row[1]}</td>
    <td id="copyid" class="column-5 text-center">${row[2]}</td>
    <td class="column-5 text-center">${row[4]}</td>
</tr>`;
                }
            });
        } else {
            resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">ไม่พบข้อมูล</td></tr>`;
        }
    } catch (error) {
        console.error('Error loading data:', error);
        alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
    }
}



async function filterRecords() {
    const searchValue = document.getElementById('datasearchname').value.toLowerCase();
    const url = `https://script.google.com/macros/s/AKfycbyJ0LNyJ97NiI4ooMEZZItKh-GXy9dBy_zREkx4OwJ5hFF4Ubzl_3qml_l--wfagRg/exec?searchValue=${searchValue}`;

    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();
        const resultDiv1 = document.getElementById('searchdataload');
        resultDiv1.innerHTML = ''; // เคลียร์ข้อมูลเก่าออก

        if (data.length > 0) {
            let found = false; // ใช้สำหรับตรวจสอบว่ามีข้อมูลที่ตรงกับการค้นหาหรือไม่

            data.forEach(row => {
                if (row && row[1] && row[1].toLowerCase().includes(searchValue)) {
                    found = true; // พบข้อมูลที่ตรงกับการค้นหา
                    resultDiv1.innerHTML += `
                        <tr onclick="copyId(this)" style="cursor: pointer;">
                            <th class="column-5 text-center">${row[0]}</th>
                            <td class="column-6">${row[1]}</td>
                            <td class="idCardCell column-5 text-center">${row[2]}</td>
                            <td class="column-5 text-center">${row[4]}</td>
                        </tr>
                    `;
                }
            });

            // หากไม่พบข้อมูลที่ตรงกับการค้นหา
            if (!found) {
                resultDiv1.innerHTML = `<tr><td colspan="4" class="text-center">ไม่พบข้อมูลที่ตรงกับการค้นหา</td></tr>`;
            }
        } else {
            // กรณีไม่มีข้อมูลใน Google Sheets
            resultDiv1.innerHTML = `<tr><td colspan="4" class="text-center">ไม่พบข้อมูล</td></tr>`;
        }
    } catch (error) {
        console.error('Error loading records:', error);
        alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
    }
}


async function displayNextNumber() {
    const url = `https://script.google.com/macros/s/AKfycbx5iWbvWY0xXDxW6N38l2LkIM4jDrGcjpvlk9Y9UYEwJfeIDQLSH1VCZnp7hCtHmi5N/exec`; // ใส่ URL ของ Web App

    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error("Network response was not ok " + response.statusText);
        }

        const data = await response.json(); // แปลง response เป็น JSON
        const nextNumber = data.nextNumber; // ดึงค่าที่ส่งมาจาก backend

        // แสดงผลในหน้า HTML
        const newNumberElement = document.getElementById('numb');
        newNumberElement.innerText = nextNumber;

    } catch (error) {
        console.error('Error fetching data:', error);
        const newNumberElement = document.getElementById('numb');
        newNumberElement.innerText = "Error"; // แสดงข้อความเมื่อเกิดข้อผิดพลาด
    }
}


async function getNextNumber() {
    const url = `https://script.google.com/macros/s/AKfycbxjea3aR5AYbPn90Xfh2l670mxa0yzNo2va1RzBMRYykrgVjri37x4sBTatmztQQWPQ/exec`; // เปลี่ยนเป็น URL ของ Web App

    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error("Network response was not ok " + response.statusText);
        }
        const data = await response; // ข้อมูลที่รับจาก Web App ในรูปแบบ JSON

        if (data && data.length > 0) {
            const lastNumber = Math.floor(data); ; // สมมติว่าเลขอยู่ในคอลัมน์แรก
            const newNumber = document.getElementById('numb');
            const nextNumber = Number(lastNumber);
            console.log(nextNumber);
            newNumber.innerText = parseInt(nextNumber) + 1; // เพิ่ม 1 ไปยังเลขล่าสุด
        } else {
            const newNumber = document.getElementById('numb');
            newNumber.innerText = 1; // ถ้าไม่มีข้อมูลใน Google Sheets ให้เริ่มที่ 1
        }
    } catch (error) {
        console.error('Error fetching data:', error);
    }
}


function searchProgram() {
    const apiKey = 'AIzaSyCUkklmX-ewNdmcDfJCb5FImBwfN0F4wjg';
    const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
    const angesheet2 = 'program!A2:Z';

    const url1 = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet2}?key=${apiKey}`;
    showLoading();
    fetch(url1)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            const programdetail = document.getElementById('programdetail');
            const programName = document.getElementById('program').textContent;
            programdetail.textContent = ""; // ล้างข้อมูลก่อนหน้า

            let found = false; // กำหนดค่า found เป็น false เริ่มต้น

            // ค้นหาโปรแกรมที่ตรงกับชื่อที่ให้มา
            if (data.values && data.values.length > 0) { // ตรวจสอบว่ามีข้อมูลใน data.values
                data.values.forEach(row => {
                    if (row[0] === programName) {
                        programdetail.innerHTML += `<p>- ${row[1]}</p>`; // แสดงรายละเอียดโปรแกรม
                        found = true; // เปลี่ยนค่า found เป็น true ถ้าพบโปรแกรม
                    }
                });
                hideLoading();
            }

            // หากไม่พบโปรแกรมที่ค้นหา
            if (!found) {
                programdetail.innerHTML = `<p>ไม่พบโปรแกรมที่ต้องการ</p>`;
            }
        })
        .catch(error => {
            console.error('Error:', error);
            const programdetail = document.getElementById('programdetail');
            programdetail.innerHTML = `<p>เกิดข้อผิดพลาดในการดึงข้อมูล</p>`;
        });
}


function printSticker() {
    
    $(".modalrecieve").css('display', 'block');
}

function closeSticker() {
    $(".modalrecieve").css('display', 'none');
}

function printSpecimen() {
    printResult();
    closeSticker();
    clearPage();
    displayNextNumber();

}

function printResult() {
    const modalContent = document.querySelector('#sticker').innerHTML;
    const originalContent = document.body.innerHTML;

    // แสดงเฉพาะเนื้อหาที่ต้องการพิมพ์
    document.body.innerHTML = modalContent;
    window.print();
    document.body.innerHTML = originalContent;
}





function clearPage() {
    var searchKey = document.getElementById('searchKey');
    var regisid = document.getElementById('registernumber');
    var name = document.getElementById('name');
    var idcard = document.getElementById('idcard');
    var age = document.getElementById('age');
    var birthday = document.getElementById('birthday');
    var program = document.getElementById('program');
    var programDetail = document.getElementById('programdetail');


    searchKey.innerHTML = '';
    regisid.textContent = '';
    name.textContent = '';
    idcard.textContent = '';
    age.textContent = '';
    birthday.textContent = '';
    program.textContent = '';
    programDetail.textContent = '';

}







function formatDateTime() {
    const now = new Date(); // ดึงวันที่และเวลาปัจจุบัน

    // ดึงค่าวัน, เดือน, ปี
    const day = String(now.getDate()).padStart(2, '0'); // เติมเลข 0 ถ้าวันน้อยกว่า 10
    const month = String(now.getMonth() + 1).padStart(2, '0'); // เดือนใน JavaScript เริ่มที่ 0
    const year = String(now.getFullYear()).slice(2); // เอาสองหลักสุดท้ายของปี

    // ดึงค่า ชั่วโมงและนาที
    const hours = String(now.getHours()).padStart(2, '0'); // เติมเลข 0 ถ้าชั่วโมงน้อยกว่า 10
    const minutes = String(now.getMinutes()).padStart(2, '0'); // เติมเลข 0 ถ้านาทีน้อยกว่า 10

    // รวมเป็นรูปแบบที่ต้องการ: dd/mm/yy, hh:mm
    return `${day}/${month}/${year}, ${hours}:${minutes}`;
}

// แสดงวันที่และเวลาปัจจุบัน
console.log(formatDateTime());

function updateDateTime() {
    const now = new Date();
    const date = now.toLocaleDateString('th-TH');
    const time = now.toLocaleTimeString('th-TH', { hour12: false }); // ใช้ 24 ชั่วโมง
    document.getElementById('datetime').textContent = `${date} ${time}`;
}

// เรียกใช้ updateDateTime ทุก 1 วินาที
setInterval(updateDateTime, 1000);

function formatDateTime(dateTime) {
    const date = new Date(dateTime);

    // ดึงข้อมูลส่วนของวัน, เดือน, ปี, ชั่วโมง และนาที
    const day = String(date.getDate()).padStart(2, '0'); // เติม 0 ถ้าตัวเลขน้อยกว่า 10
    const month = String(date.getMonth() + 1).padStart(2, '0'); // เดือนเริ่มจาก 0, ต้อง +1
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0'); // ชั่วโมง
    const minutes = String(date.getMinutes()).padStart(2, '0'); // นาที

    // รวมผลลัพธ์เป็นรูปแบบที่ต้องการ
    return `${day}/${month}/${year},${hours}:${minutes}`;
}


function birthdayFormat(dateTime) {
    const date = new Date(dateTime);

    // ดึงข้อมูลส่วนของวัน, เดือน, ปี, ชั่วโมง และนาที
    const day = String(date.getDate()).padStart(2, '0'); // เติม 0 ถ้าตัวเลขน้อยกว่า 10
    const month = String(date.getMonth() + 1).padStart(2, '0'); // เดือนเริ่มจาก 0, ต้อง +1
    const year = date.getFullYear();
   

    // รวมผลลัพธ์เป็นรูปแบบที่ต้องการ
    return `${day}/${month}/${year}`;
}


function updateDate() {
    const now = new Date();
    const date = now.toLocaleDateString('th-TH');
    const time = now.toLocaleTimeString('th-TH');
    const current = `${date}`;
    return current
}
function openSheet() {
    $('#sheet').css('display', 'flex');
}

function closeSheet() {
    $('#sheet').css('display', 'none');
}

function getAlert() {
    $('#alert').css('display', 'flex');
}

function closeAlert() {
    $('#alert').css('display', 'none');
}

function closeNewRegister() {
    var n1 = document.getElementById('idcard');
    var n2 = document.getElementById('newidcard').value;
    n1.innerText = n2;
    $('.modalregister').css('display', 'none');
}

function openNewRegister() {
    clearPage();
    clearRegisterPage();
    $('.modalregister').css('display', 'block');
}

function openSpecimen() {
    $('.modalspecimen').css('display', 'flex');
}

function closeSpecimen() {
    $('.modalspecimen').css('display', 'none');
}














function calculateAge() {
    var birthdate = document.getElementById('birthdate').value;

    if (birthdate) {
        // แปลงวันที่เป็น Date object
        var birthDateObj = new Date(birthdate);
        var today = new Date(); // วันที่ปัจจุบัน
        var age = today.getFullYear() - birthDateObj.getFullYear(); // คำนวณอายุคร่าว ๆ

        // ตรวจสอบเดือนและวัน เพื่อให้แน่ใจว่าคำนวณอายุถูกต้อง (เผื่อยังไม่ถึงวันเกิดปีนี้)
        var monthDiff = today.getMonth() - birthDateObj.getMonth();
        if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDateObj.getDate())) {
            age--; // ถ้าเดือนนี้ยังไม่ถึงวันเกิด ให้ลบอายุออก 1 ปี
        }

        // แสดงผลใน <span> ที่มี id="newage"
        document.getElementById('newage').textContent = age;
    } else {
        // ถ้าไม่มีวันเดือนปีเกิด ให้ล้างค่าใน <span>
        document.getElementById('newage').textContent = '';
    }
}





function clearSpecimen() {

    var barcode = document.getElementById('inputbar');
    var bid = document.getElementById('barregisterid');
    var bname = document.getElementById('barname');
    barcode.value = '';
    bid.innerText = '';
    bname.innerText = '';
}

function getCurrentDateTime() {
    const now = new Date();
    return now.toLocaleString(); // Date and time format
}

function updateDateTime() {
    const now = new Date();
    const date = now.toLocaleDateString('th-TH');
    const time = now.toLocaleTimeString('th-TH');
    document.getElementById('datetime').innerHTML = `${date} ${time}`;
}


setInterval(updateDateTime, 1000);

function openSheet() {
    $(".modalsheet").css('display', 'block');
}
function closeRegist() {
    clearRegisterPage();
    $(".modalregister").css('display', 'none');
}
function closeSheet() {
    $(".modalsheet").css('display', 'none');
}



function closeSearch() {
    $(".modalsearch").css('display', 'none');
}



function printResult() {
    var modalContent = document.querySelector('#sticker').innerHTML;
    var originalContent = document.body.innerHTML;


    document.body.innerHTML = modalContent;
    window.print();
    document.body.innerHTML = originalContent;



}


function clearPage() {
    var searchKey = document.getElementById('searchKey');
    var regisid = document.getElementById('registernumber');
    var name = document.getElementById('name');
    var idcard = document.getElementById('idcard');
    var age = document.getElementById('age');
    var birthday = document.getElementById('birthday');
    var program = document.getElementById('program');
    var programDetail = document.getElementById('programdetail');
    var card = document.getElementById('card');
    var depart = document.getElementById('depart');
    searchKey.value = '';
    regisid.textContent = '';
    name.textContent = '';
    idcard.textContent = '';
    age.textContent = '';
    birthday.textContent = '';
    program.textContent = '';
    programDetail.textContent = '';
    card.textContent = '';
    depart.textContent = '';
}


function clearRegisterPage() {

    var newregisid = document.getElementById('newid');
    var newname = document.getElementById('newname');
    var newidcard = document.getElementById('newidcard');
    var newage = document.getElementById('newage');
    var newbirthday = document.getElementById('birthdate');
    var newprogram = document.getElementById('newprogram');
    var newcard = document.getElementById('newcard');
    var newdepart = document.getElementById('newdepart');

    newdepart.value = '';
    newregisid.value = '';
    newname.value = '';
    newidcard.value = '';
    newage.textContent = '';
    newbirthday.textContent = '';
    newprogram.value = '';
    newcard.value = '';
}


function openSearch() {
    $(".modalsearch").css('display', 'block');
     clearPage();
  
}

function closeSearch() {
    $(".modalsearch").css('display', 'none');
}


function copyId(element) {
    // element คือแถวที่ถูกคลิก
    const idcard = element.querySelector(".idCardCell").textContent.trim();
    const copyend = document.getElementById('idcard');

    if (copyend) {
        copyend.textContent = idcard; 
    }

    closeSearch();
    searchDataFromId();
}







function deleteAllFilter(){
     loadAllRecordsSearch();
     var search1 = document.getElementById('datasearchid');
     var search2 = document.getElementById('datasearchname');
    search1.value = '';
    search2.value = '';
}


window.onload = function () {
  loadAllRecords();
  displayNextNumber(); 
  updateDateTime();
  loadDataTable();
  loadAllCount();

}


