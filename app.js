const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

// Data dinamis untuk tabel
let students = [
    { name: "John Doe", age: 16, attendance: "90%" },
    { name: "Jane Smith", age: 15, attendance: "85%" },
    { name: "Michael Brown", age: 17, attendance: "95%" },
];

// next data dari hasil query ke database atau dari collection postman

// Fungsi untuk memformat tanggal dalam format DD-MM-YYYY-HH-mm-ss
function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Bulan dimulai dari 0
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');

    // perhatikan penamaan file di setiap OS
    return `${day}${month}${year}${hours}${minutes}${seconds}`;
}

// Fungsi untuk membuat laporan
function generateReport(studentsData) {
    try {
        // Membaca template DOCX
        const templatePath = path.join(__dirname, "templates", "report-template.docx");
        const templateContent = fs.readFileSync(templatePath);
        const zip = new PizZip(templateContent);

        // Menghubungkan dengan Docxtemplater
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Data untuk diisi ke dalam template
        const data = { students: studentsData };

        // Render data ke template
        doc.render(data);

        // Mendapatkan tanggal sekarang
        const currentDate = new Date();

        // Format tanggal dengan jam, menit, dan detik
        const formattedDate = formatDate(currentDate);

        // Menyimpan hasil ke file baru
        const outputPath = path.join(__dirname, "outputs", `report-${formattedDate}.docx`);
        const buffer = doc.getZip().generate({ type: "nodebuffer" });
        fs.writeFileSync(outputPath, buffer);

        console.log(`Laporan berhasil dibuat di: ${outputPath}`);
    } catch (err) {
        console.error("Error saat membuat laporan:", err);
    }
}

generateReport(students);
