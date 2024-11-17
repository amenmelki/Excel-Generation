
const fs = require('fs');
const path = require('path');
const xlsx = require("xlsx"); 
const ExcelJS = require('exceljs');
const initializeWorkbook = async (filePath) => {
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    } else {
        workbook.addWorksheet('Example1');
    }

    return workbook;
};

const addDataToWorksheet = (worksheet, headers, data) => {
    if (worksheet.rowCount === 0) {
        worksheet.addRow(headers);

        const headerRow = worksheet.getRow(1);
        headerRow.font = { 
            name: 'Candara Light',
            size: 12,
            bold: true,
            color: { argb: '006400' } 
        };
        headerRow.commit();
    }

    data.forEach((row) => {
        worksheet.addRow(row);
    });

    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: '000000' } },
                left: { style: 'thin', color: { argb: '000000' } },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: { style: 'thin', color: { argb: '000000' } }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
    });

    worksheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            maxLength = Math.max(maxLength, cellLength);
        });
        column.width = maxLength + 2;
    });
};

exports.testexcel = async (req, res) => {
    try {
        const {
            nb_inscri,
            cin,
            cin_date,
            candidateTripleName,
            candidateLastName,
            birthDate,
            birthPlace,
            gender,
            address,
            phone,
            state,
            postalCode,
            email,
            specialization,
            fileNumber,
            obtainedDegree,
            bachelorAverage,
            graduationAverage,
            total_point
        } = req.body;

        const candidateFolder = path.join(__dirname, `../uploads/testexcel`);
        if (!fs.existsSync(candidateFolder)) {
            fs.mkdirSync(candidateFolder, { recursive: true });
        }

        const excelFilePath = path.join(candidateFolder, `testexcel.xlsx`);

        const workbook = await initializeWorkbook(excelFilePath);
        const worksheet = workbook.getWorksheet('Example1');
        const headers1 = [
            "رقم التسجيل",
            "رقم بطاقة التعريف الوطنية",
            "تاريخ الإصدار ",
            "الأسم الثلاثي للمترشح",
            "اللقب",
            "تاريخ الولادة ",
            "مكان الولادة ",
            "الجنس",
            "عنوان المترشح بكل دقة",
            "رقم الهاتف",
            "الولاية",
            "الترقيم البريدي",
            "البريد الالكتروني",
            "الخطة المترشح لها",
            "الصنف",
            "الشهادة العلمية المتحصل عليها",
            "أختصاص شهادة الكفاءة المهنية",
            "سنة الحصول على الشهادة العلمية ",
            "مجموع النقاط"
        ];
        const data1 = [
            [nb_inscri, cin, cin_date, candidateTripleName, candidateLastName, birthDate, birthPlace, gender, address, phone, state, postalCode, email, specialization, fileNumber, obtainedDegree, bachelorAverage, graduationAverage, total_point]
        ];

        addDataToWorksheet(worksheet, headers1, data1); 
        let worksheet2 = workbook.getWorksheet('Example2');
        if (!worksheet2) {
            worksheet2 = workbook.addWorksheet('Example2 ');
        }

        const headers2 = [
            "رقم التسجيل",
            "رقم بطاقة التعريف الوطنية",
            "تاريخ الإصدار ",
            "الأسم الثلاثي للمترشح",
            "اللقب",
            "تاريخ الولادة ",
            "مكان الولادة ",
            "الجنس",
            "عنوان المترشح بكل دقة",
            "رقم الهاتف",
            "الولاية",
            "الترقيم البريدي",
            "البريد الالكتروني",
            "الخطة المترشح لها",
            "الصنف",
            "مجموع النقاط"
        ];
        const data2 = [
            [
                nb_inscri,       
                cin,             
                cin_date,        
                candidateTripleName,
                candidateLastName, 
                birthDate,       
                birthPlace,      
                gender,          
                address,         
                phone,           
                state,          
                postalCode,       
                email,           
                specialization, 
                fileNumber,      
                total_point       
            ]
        ];

    addDataToWorksheet(worksheet2, headers2, data2);

        await workbook.xlsx.writeFile(excelFilePath);

        return res.status(201).json({
            message: "Data added to Excel file successfully with bold headers, black borders, center alignment, and auto-resized columns.",
        });
    } catch (err) {
        console.log("Error in testexcel: ", err);
        return res.status(400).json({
            message: "Failed to add data to Excel file.",
        });
    }
};

