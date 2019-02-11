---
title: 'Fungsi Lookup dan Referensi (Rumus VLOOKUP/HLOOKUP)'
date: Sun, 14 Jan 2018 16:00:42+07:00
draft: false
tag: [belajar excel, rumus excel hlookup, rumus excel vlookup, Tutorial Excel]
---

Fungsi pencarian atau referensi terdapat beberapa jenis 2 diantaranya yang sering dipakai, yakni VLOOKUP dan HLOOKUP. Berikut ini penjelasannya.

### VLOOKUP

Fungsi VLOOKUP (Vertical Lookup) akan mengisi sel dengan nilai yang mengacu kepada tabel vertikal yang anda tentukan. Nilai tabel tersebut ditentukan dari nilai kolom di sebelah kiri sel. Berikut ini adalah rumusnya: =VLOOKUP(lookup\_value; table\_array; col\_index\_num; \[range\_lookup\]) *\[range\_lookup\] = opsional (boleh diisi dan boleh juga tidak, nilai defaultnya adalah TRUE) **Penjelasan Rumus VLOOKUP** ![](https://www.bisnis7.com/wp-content/uploads/2018/01/tutorial-excel-rumus-vlookup.jpg) Perhatikan sel G3, isi formulanya adalah =VLOOKUP(G2;B3:D6;3). Ini artinya nilai G3 ditentukan oleh nilai G2 dimana nilai G2 harus terdapat pada tabel B3:D6 (ditandai dengan kotak merah) dan referensi nilai yang diambil adalah kolom nomor 3 (kolom harga). Mudah bukan? Apabila pada kolom G2 tidak terdapat pada tabel B3:D6, maka akan terdapat error pada nilai G3.

### VLOOKUP

Fungsi HLOOKUP (Horizontal Lookup) akan mengisi sel dengan nilai yang mengacu kepada tabel horizontal yang anda tentukan. Nilai tabel tersebut ditentukan dari nilai kolom di sebelah kiri sel.