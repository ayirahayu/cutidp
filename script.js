const cutiList = JSON.parse(localStorage.getItem("cutiList")) || [];
let editIndex = -1; // Untuk mengetahui apakah sedang dalam mode edit

function updateTable() {
    const tbody = document.getElementById("cutiTable").querySelector("tbody");
    tbody.innerHTML = "";
    cutiList.forEach((cuti, index) => {
        const row = tbody.insertRow();
        row.insertCell().textContent = cuti.nama;
        row.insertCell().textContent = cuti.grup;
        row.insertCell().textContent = cuti.tanggalMulai;
        row.insertCell().textContent = cuti.tanggalSelesai;
        row.insertCell().textContent = cuti.alasan;
        row.insertCell().textContent = cuti.pengganti;

        // Tambahkan tombol Edit dan Delete
        const actionsCell = row.insertCell();
        const editButton = document.createElement("button");
        editButton.textContent = "Edit";
        editButton.onclick = () => editCuti(index);

        const deleteButton = document.createElement("button");
        deleteButton.textContent = "Delete";
        deleteButton.onclick = () => deleteCuti(index);

        actionsCell.appendChild(editButton);
        actionsCell.appendChild(deleteButton);
    });
}

function saveToLocalStorage() {
    localStorage.setItem("cutiList", JSON.stringify(cutiList));
}

document.getElementById("cutiForm").addEventListener("submit", function (e) {
    e.preventDefault();

    const nama = document.getElementById("nama").value;
    const grup = document.getElementById("grup").value;
    const tanggalMulai = document.getElementById("tanggalMulai").value;
    const tanggalSelesai = document.getElementById("tanggalSelesai").value;
    const alasan = document.getElementById("alasan").value;
    const pengganti = document.getElementById("pengganti").value;

    const cuti = { nama, grup, tanggalMulai, tanggalSelesai, alasan, pengganti };

    if (editIndex === -1) {
        // Jika tidak dalam mode edit, tambahkan data baru
        cutiList.push(cuti);
    } else {
        // Jika dalam mode edit, perbarui data
        cutiList[editIndex] = cuti;
        editIndex = -1;
        document.getElementById("cutiForm").querySelector("button[type='submit']").textContent = "Tambah Cuti";
    }

    updateTable();
    saveToLocalStorage();
    document.getElementById("cutiForm").reset();
});

function editCuti(index) {
    const cuti = cutiList[index];
    document.getElementById("nama").value = cuti.nama;
    document.getElementById("grup").value = cuti.grup;
    document.getElementById("tanggalMulai").value = cuti.tanggalMulai;
    document.getElementById("tanggalSelesai").value = cuti.tanggalSelesai;
    document.getElementById("alasan").value = cuti.alasan;
    document.getElementById("pengganti").value = cuti.pengganti;

    editIndex = index;
    document.getElementById("cutiForm").querySelector("button[type='submit']").textContent = "Simpan Perubahan";
}

function deleteCuti(index) {
    cutiList.splice(index, 1);
    updateTable();
    saveToLocalStorage();
}

updateTable();


function exportToExcel() {
    const data = [
        ['Nama Karyawan', 'Grup', 'Tanggal Mulai', 'Tanggal Selesai', 'Alasan Cuti', 'Pengganti']
    ];
    
    cutiList.forEach(cuti => 
        data.push([cuti.nama, cuti.grup, cuti.tanggalMulai, cuti.tanggalSelesai, cuti.alasan, cuti.pengganti])
    );

    // Membuat worksheet dan workbook
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daftar Cuti");

    // Menambahkan style ke header dengan perpaduan warna tema BTPN Syariah
    const headerStyle = { 
        font: { bold: true, color: { rgb: "FFFFFF" } }, // Warna font putih
        fill: { fgColor: { rgb: "FFC000" } }, // Warna kuning (perpaduan tema BTPN Syariah)
        alignment: { horizontal: "center", vertical: "center" }
    };

    // Terapkan style ke header
    const headerRange = XLSX.utils.decode_range(ws['!ref']);
    for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {
        const headerCell = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!ws[headerCell]) continue;
        ws[headerCell].s = headerStyle;
    }

    // Memberikan border ke seluruh sel
    const borderStyle = {
        top: { style: "thin", color: { rgb: "003b71" } }, // warna biru tua
        bottom: { style: "thin", color: { rgb: "003b71" } },
        left: { style: "thin", color: { rgb: "003b71" } },
        right: { style: "thin", color: { rgb: "003b71" } }
    };

    for (let R = headerRange.s.r; R <= headerRange.e.r; ++R) {
        for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {
            const cell = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cell]) continue;
            ws[cell].s = ws[cell].s || {};
            ws[cell].s.border = borderStyle;
        }
    }

    // Menyesuaikan lebar kolom otomatis agar rapi
    const colWidths = data[0].map((_, i) => ({
        wch: Math.max(
            ...data.map(row => row[i] ? row[i].toString().length + 2 : 12) // lebar kolom
        )
    }));
    ws['!cols'] = colWidths;

    // Menyimpan file Excel dengan nama yang diinginkan
    XLSX.writeFile(wb, "List_Cuti.xlsx");
}

