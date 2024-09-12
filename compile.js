const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Tentukan path folder 'data' dan file output
const dataFolderPath = path.join(__dirname, 'data');
const outputFilePath = path.join(__dirname, 'combined.xlsx');

// Fungsi untuk membaca file JSON dan menggabungkannya
function mergeJsonFilesToExcel(directoryPath, outputFilePath) {
    fs.readdir(directoryPath, (err, files) => {
        if (err) {
            console.error('Gagal membaca folder:', err);
            return;
        }

        // Filter hanya file dengan ekstensi .json
        const jsonFiles = files.filter(file => file.endsWith('.json'));

        let combinedData = [];

        // Baca setiap file JSON dan gabungkan datanya
        jsonFiles.forEach((file, index) => {
            const filePath = path.join(directoryPath, file);
            const fileData = fs.readFileSync(filePath, 'utf8');
            try {
                const jsonData = JSON.parse(fileData);
                combinedData = combinedData.concat(jsonData);
            } catch (parseError) {
                console.error(`Gagal mem-parsing file ${file}:`, parseError);
            }

            // Jika sudah membaca semua file, simpan hasilnya ke file output
            if (index === jsonFiles.length - 1) {
                // Konversi data gabungan menjadi sheet Excel
                const ws = XLSX.utils.json_to_sheet(combinedData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Combined Data');

                // Tulis workbook ke file output
                XLSX.writeFile(wb, outputFilePath);

                console.log(`Berhasil menggabungkan file JSON menjadi ${outputFilePath}`);
            }
        });
    });
}

// Panggil fungsi untuk menggabungkan file JSON dan menyimpannya sebagai Excel
mergeJsonFilesToExcel(dataFolderPath, outputFilePath);
