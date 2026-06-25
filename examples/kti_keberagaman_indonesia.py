"""
KTI Lengkap: Keberagaman Indonesia
Generate dokumen KTI sesuai standar akademik Indonesia
"""
from docx_idn import KTI

doc = KTI(
    # ============================================
    # IDENTITAS
    # ============================================
    judul="KEBERAGAMAN BUDAYA INDONESIA: STUDI TENTANG PELESTARIAN DAN PEKEMBANGAN NILAI-NILAI TRADISIONAL DI ERA GLOBALISASI",
    judul_en="CULTURAL DIVERSITY OF INDONESIA: A STUDY ON PRESERVATION AND DEVELOPMENT OF TRADITIONAL VALUES IN THE GLOBALIZATION ERA",
    penulis="Rina Wulandari",
    nim="2024010045",
    institusi="UNIVERSITAS NEGERI JAKARTA",
    fakultas="Ilmu Sosial dan Ilmu Politik",
    prodi="Sosiologi",
    angkatan="2024",
    pembimbing="Dr. H. Bambang Sutrisno, M.Si.",
    tanggal="Juni 2026",

    # ============================================
    # ABSTRAK
    # ============================================
    abstrak_id=(
        "Keberagaman merupakan salah satu karakteristik utama bangsa Indonesia. "
        "Sebagai negara kepulauan terbesar di dunia dengan lebih dari 17.000 pulau, "
        "300 lebih suku bangsa, dan 700 lebih bahasa daerah, Indonesia menyimpan "
        "kekayaan budaya yang luar biasa. Penelitian ini bertujuan untuk menganalisis "
        "faktor-faktor yang mempengaruhi pelestarian keberagaman budaya Indonesia serta "
        "strategi pengembangan nilai-nilai tradisional di era globalisasi. Metode "
        "penelitian yang digunakan adalah kualitatif dengan pendekatan deskriptif. "
        "Data dikumpulkan melalui wawancara mendalam, observasi partisipan, dan "
        "studi dokumentasi. Hasil penelitian menunjukkan bahwa pelestarian keberagaman "
        "budaya Indonesia menghadapi tantangan serius berupa urbanisasi, pengaruh "
        "budaya asing, dan minimnya regenerasi. Namun, berbagai upaya pelestarian "
        "telah dilakukan melalui pendidikan budaya, festival kesenian, dan pemanfaatan "
        "teknologi digital. Kesimpulannya, pelestarian keberagaman budaya memerlukan "
        "kolaborasi antara pemerintah, masyarakat, dan sektor swasta untuk memastikan "
        "keberlanjutan nilai-nilai tradisional bangsa Indonesia."
    ),
    abstrak_en=(
        "Diversity is one of the main characteristics of the Indonesian nation. "
        "As the largest archipelago in the world with more than 17,000 islands, "
        "over 300 ethnic groups, and more than 700 regional languages, Indonesia "
        "holds extraordinary cultural richness. This study aims to analyze the "
        "factors influencing the preservation of Indonesian cultural diversity and "
        "strategies for developing traditional values in the globalization era. "
        "The research method used is qualitative with a descriptive approach. "
        "Data were collected through in-depth interviews, participant observation, "
        "and documentation study. The results show that the preservation of "
        "Indonesian cultural diversity faces serious challenges in the form of "
        "urbanization, foreign cultural influence, and minimal regeneration. "
        "However, various preservation efforts have been carried out through "
        "cultural education, arts festivals, and the use of digital technology. "
        "In conclusion, the preservation of cultural diversity requires collaboration "
        "between the government, community, and private sector to ensure the "
        "sustainability of Indonesia's traditional values."
    ),

    # ============================================
    # KATA PENGANTAR
    # ============================================
    kata_pengantar=(
        "Puji syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa karena "
        "berkat rahmat dan karunia-Nya, kami dapat menyelesaikan Karya Tulis "
        "Ilmiah yang berjudul \"Keberagaman Budaya Indonesia: Studi Tentang "
        "Pelestarian dan Pengembangan Nilai-Nilai Tradisional di Era Globalisasi\".\n\n"
        "KTI ini disusun untuk memenuhi salah satu syarat dalam menyelesaikan "
        "program studi S1 Sosiologi di Universitas Negeri Jakarta. Dalam "
        "penyusunan KTI ini, kami mendapatkan banyak bantuan dan dukungan "
        "dari berbagai pihak.\n\n"
        "Kami mengucapkan terima kasih kepada:\n"
        "1. Dr. H. Bambang Sutrisno, M.Si. selaku dosen pembimbing yang telah "
        "memberikan bimbingan dan arahan selama penyusunan KTI ini.\n"
        "2. Seluruh dosen Program Studi Sosiologi Universitas Negeri Jakarta "
        "yang telah memberikan ilmu dan pengetahuan selama perkuliahan.\n"
        "3. Keluarga besar kami yang selalu memberikan doa dan dukungan.\n"
        "4. Semua pihak yang tidak dapat kami sebutkan satu per satu.\n\n"
        "Kami menyadari bahwa KTI ini masih jauh dari sempurna. Oleh karena "
        "itu, kami mengharapkan kritik dan saran yang membangun dari para "
        "pembaca. Semoga KTI ini dapat memberikan manfaat bagi perkembangan "
        "ilmu pengetahuan, khususnya di bidang sosiologi budaya.\n\n"
        "Jakarta, Juni 2026\n"
        "Penulis"
    ),

    # ============================================
    # BAB-BAB
    # ============================================
    bab=[
        # BAB I: Pendahuluan
        {
            "judul": "PENDAHULUAN",
            "sections": [
                {
                    "judul": "Latar Belakang",
                    "isi": [
                        "Indonesia merupakan negara yang sangat unik di dunia. "
                        "Sebagai negara kepulauan terbesar dengan luas wilayah "
                        "1.904.569 km², Indonesia terdiri dari 17.008 pulau "
                        "yang membentang dari Sabang sampai Merauke. Keunikan "
                        "ini tidak hanya terlihat dari aspek geografis, tetapi "
                        "juga dari aspek budaya, bahasa, dan adat istiadat.",

                        "Berdasarkan data dari Badan Pusat Statistik (BPS) tahun "
                        "2024, Indonesia memiliki lebih dari 300 suku bangsa yang "
                        "mendiami berbagai daerah. Setiap suku bangsa memiliki "
                        "kebudayaan yang khas dan berbeda satu sama lain. Bahkan, "
                        "menurut data Kementerian Pendidikan dan Kebudayaan, "
                        "terdapat lebih dari 700 bahasa daerah yang digunakan "
                        "di seluruh wilayah Indonesia.",

                        "Keberagaman ini merupakan aset bangsa yang sangat "
                        "berharga. Bung Karno, proklamator kemerdekaan Indonesia, "
                        "pernah menyatakan \"Bhinneka Tunggal Ika\" yang berarti "
                        "Berbeda-beda tetapi tetap satu. Semboyan ini menjadi "
                        "filosofi dasar kehidupan berbangsa dan bernegara di "
                        "Indonesia.",

                        "Namun, di era globalisasi seperti saat ini, keberagaman "
                        "budaya Indonesia menghadapi tantangan yang serius. "
                        "Pengaruh budaya asing yang masuk melalui media sosial "
                        "dan teknologi informasi menyebabkan pergeseran nilai "
                        "di kalangan generasi muda. Banyak tradisi dan adat "
                        "istiadat yang mulai ditinggalkan oleh generasi penerus.",

                        "Menurut survei yang dilakukan oleh UNESCO tahun 2023, "
                        "sekitar 40% warisan budaya Indonesia berada dalam "
                        "kondisi terancam punah. Angka ini sangat mengkhawatirkan "
                        "dan membutuhkan perhatian serius dari semua pihak. "
                        "Oleh karena itu, penelitian ini bertujuan untuk menganalisis "
                        "strategi pelestarian dan pengembangan keberagaman budaya "
                        "Indonesia di era globalisasi."
                    ]
                },
                {
                    "judul": "Rumusan Masalah",
                    "isi": [
                        "Berdasarkan latar belakang yang telah diuraikan, "
                        "maka rumusan masalah dalam penelitian ini adalah:",
                        "1. Bagaimana kondisi keberagaman budaya Indonesia saat ini?",
                        "2. Faktor-faktor apa saja yang mempengaruhi pelestarian "
                        "keberagaman budaya Indonesia?",
                        "3. Bagaimana strategi pengembangan nilai-nilai tradisional "
                        "di era globalisasi?",
                        "4. Apa peran masyarakat dalam pelestarian keberagaman budaya?"
                    ]
                },
                {
                    "judul": "Tujuan Penelitian",
                    "isi": [
                        "Penelitian ini bertujuan untuk:",
                        "1. Mendeskripsikan kondisi keberagaman budaya Indonesia "
                        "saat ini berdasarkan data empiris.",
                        "2. Menganalisis faktor-faktor yang mempengaruhi pelestarian "
                        "keberagaman budaya Indonesia.",
                        "3. Merumuskan strategi pengembangan nilai-nilai tradisional "
                        "di era globalisasi.",
                        "4. Mengidentifikasi peran masyarakat dalam pelestarian "
                        "keberagaman budaya."
                    ]
                },
                {
                    "judul": "Manfaat Penelitian",
                    "isi": [
                        "Penelitian ini diharapkan dapat memberikan manfaat sebagai berikut:",
                        "1. Manfaat Teoritis: Memberikan kontribusi pada pengembangan "
                        "ilmu sosiologi budaya, khususnya mengenai dinamika pelestarian "
                        "keberagaman di masyarakat multikultural.",
                        "2. Manfaat Praktis: a) Bagi pemerintah: menjadi bahan "
                        "pertimbangan dalam menyusun kebijakan pelestarian budaya. "
                        "b) Bagi masyarakat: meningkatkan kesadaran akan pentingnya "
                        "melestarikan budaya lokal. c) Bagi akademisi: menjadi "
                        "referensi untuk penelitian selanjutnya."
                    ]
                }
            ]
        },

        # BAB II: Tinjauan Pustaka
        {
            "judul": "TINJAUAN PUSTAKA",
            "sections": [
                {
                    "judul": "Pengertian Keberagaman",
                    "isi": [
                        "Keberagaman atau keragaman merupakan suatu kondisi "
                        "di mana terdapat perbedaan antara satu kelompok dengan "
                        "kelompok lainnya. Dalam konteks kehidupan bermasyarakat, "
                        "keberagaman mencakup berbagai aspek seperti suku, agama, "
                        "ras, budaya, dan bahasa.",

                        "Menurut Koentjaraningrat (2015), keberagaman budaya "
                        "merupakan kondisi di mana suatu masyarakat memiliki "
                        "lebih dari satu kebudayaan yang hidup dan berkembang "
                        "secara bersamaan. Keberagaman ini terbentuk melalui "
                        "proses interaksi sosial yang panjang dan kompleks."
                    ]
                },
                {
                    "judul": "Karakteristik Keberagaman Indonesia",
                    "sections": [
                        {
                            "judul": "Keberagaman Suku Bangsa",
                            "isi": [
                                "Indonesia memiliki lebih dari 300 suku bangsa "
                                "yang tersebar di seluruh wilayah Nusantara. "
                                "Suku bangsa terbesar adalah Jawa yang mencakup "
                                "sekitar 40% penduduk Indonesia, diikuti oleh "
                                "Sunda, Melayu, Madura, dan Batak.",

                                "Setiap suku bangsa memiliki karakteristik yang "
                                "unik. Misalnya, suku Jawa dikenal dengan budaya "
                                "yang halus dan sopan santun, suku Batak dengan "
                                "budaya yang kuat dan komunal, serta suku Dayak "
                                "dengan budaya yang sangat dekat dengan alam."
                            ]
                        },
                        {
                            "judul": "Keberagaman Bahasa",
                            "isi": [
                                "Berdasarkan data dari Badan Bahasa tahun 2024, "
                                "terdapat 718 bahasa daerah di Indonesia. Jumlah "
                                "ini menjadikan Indonesia sebagai salah satu negara "
                                "dengan keberagaman bahasa terbesar di dunia.",

                                "Bahasa daerah tidak hanya sebagai alat komunikasi, "
                                "tetapi juga sebagai penentu identitas suku bangsa. "
                                "Setiap bahasa daerah memiliki kosakata, grammar, "
                                "dan filosofi yang berbeda-beda."
                            ]
                        },
                        {
                            "judul": "Keberagaman Agama",
                            "isi": [
                                "Indonesia mengakui enam agama resmi yaitu Islam, "
                                "Kristen Protestan, Katolik, Hindu, Buddha, dan "
                                "Konghucu. Keberagaman agama ini dijamin oleh "
                                "Pancasila, khususnya sila pertama: Ketuhanan "
                                "Yang Maha Esa.",

                                "Meskipun berbeda-beda, masyarakat Indonesia "
                                "umumnya hidup berdampingan secara damai. Toleransi "
                                "antar umat beragama menjadi salah satu kekuatan "
                                "bangsa Indonesia dalam menjaga keutuhan NKRI."
                            ]
                        }
                    ]
                },
                {
                    "judul": "Tantangan Pelestarian Budaya",
                    "isi": [
                        "Globalisasi membawa pengaruh yang sangat besar bagi "
                        "kehidupan masyarakat Indonesia. Pengaruh ini dapat "
                        "terlihat dari beberapa aspek berikut:",

                        "1. Pengaruh Teknologi dan Media Sosial: Media sosial "
                        "seperti Instagram, TikTok, dan YouTube menjadi jembatan "
                        "masuknya budaya asing ke Indonesia. Generasi muda lebih "
                        "terpengaruh oleh budaya pop Korea, Jepang, atau Barat "
                        "daripada budaya lokal.",

                        "2. Urbanisasi: Migrasi penduduk dari desa ke kota "
                        "menyebabkan banyak tradisi dan adat istiadat yang "
                        "mulai ditinggalkan. Generasi muda yang tinggal di kota "
                        "cenderung tidak mengenal budaya leluhur mereka.",

                        "3. Minimnya Regenerasi: Banyak kesenian dan tradisi "
                        "tradisional yang tidak memiliki penerus. Para seniman "
                        "dan budayawan senior kesulitan menemukan generasi muda "
                        "yang berminat untuk meneruskan tradisi.",

                        "4. Komodifikasi Budaya: Beberapa pihak memanfaatkan "
                        "budaya hanya untuk kepentingan komersial tanpa memperhatikan "
                        "nilai-nilai filosofis yang terkandung di dalamnya."
                    ]
                }
            ]
        },

        # BAB III: Metodologi
        {
            "judul": "METODOLOGI PENELITIAN",
            "sections": [
                {
                    "judul": "Desain Penelitian",
                    "isi": [
                        "Penelitian ini menggunakan pendekatan kualitatif dengan "
                        "metode deskriptif. Pemilihan pendekatan kualitatif "
                        "didasarkan pada pertimbangan bahwa penelitian ini "
                        "bertujuan untuk memahami fenomena sosial secara mendalam, "
                        "termasuk makna, motif, dan tujuan dari pelestarian "
                        "keberagaman budaya.",

                        "Desain penelitian yang digunakan adalah studi kasus "
                        "multi lokasi. Peneliti memilih tiga lokasi yang "
                        "mewakili keberagaman budaya Indonesia, yaitu: "
                        "Yogyakarta (Jawa), Tanah Toraja (Sulawesi Selatan), "
                        "dan Pontianak (Kalimantan Barat)."
                    ]
                },
                {
                    "judul": "Subjek Penelitian",
                    "isi": [
                        "Subjek penelitian ini terdiri dari:",
                        "1. Tokoh adat dan budayawan di setiap lokasi penelitian "
                        "(6 orang).",
                        "2. Pemerintah daerah yang bertanggung jawab atas "
                        "pelestarian budaya (3 orang).",
                        "3. Masyarakat umum yang masih melestarikan tradisi "
                        "(18 orang).",
                        "4. Akademisi yang meneliti tentang budaya Indonesia "
                        "(3 orang).",

                        "Total subjek penelitian adalah 30 orang yang dipilih "
                        "secara purposive sampling berdasarkan kriteria tertentu."
                    ]
                },
                {
                    "judul": "Teknik Pengumpulan Data",
                    "isi": [
                        "Teknik pengumpulan data yang digunakan dalam penelitian "
                        "ini meliputi:",
                        "1. Wawancara Mendalam (In-depth Interview): Wawancara "
                        "dilakukan secara langsung dengan subjek penelitian "
                        "menggunakan panduan wawancara semi-terstruktur.",
                        "2. Observasi Partisipan: Peneliti terlibat langsung "
                        "dalam kegiatan pelestarian budaya di lokasi penelitian.",
                        "3. Studi Dokumentasi: Pengumpulan data sekunder dari "
                        "dokumen-dokumen resmi, jurnal, dan sumber literatur "
                        "lainnya."
                    ]
                },
                {
                    "judul": "Teknik Analisis Data",
                    "isi": [
                        "Analisis data menggunakan model interaktif dari Miles "
                        "dan Huberman (2014) yang terdiri dari empat komponen:",
                        "1. Pengumpulan data: proses pengumpulan data di lapangan.",
                        "2. Reduksi data: proses pemilihan, pemfokusan, dan "
                        "penyederhanaan data.",
                        "3. Penyajian data: proses pengorganisasian data agar "
                        "dapat dipahami.",
                        "4. Penarikan kesimpulan: proses verifikasi dan "
                        "pengambilan kesimpulan.",

                        "Keempat komponen ini dilakukan secara simultan dan "
                        "berkelanjutan selama proses penelitian berlangsung."
                    ]
                }
            ]
        },

        # BAB IV: Hasil dan Pembahasan
        {
            "judul": "HASIL DAN PEMBAHASAN",
            "sections": [
                {
                    "judul": "Kondisi Keberagaman Budaya Indonesia",
                    "isi": [
                        "Berdasarkan hasil penelitian di lapangan, diperoleh "
                        "gambaran bahwa keberagaman budaya Indonesia masih "
                        "cukup kuat, namun mengalami erosi di beberapa aspek. "
                        "Di Yogyakarta, misalnya, tradisi-upacara adat seperti "
                        "Bersih Desa dan Grebeg masih dilaksanakan secara "
                        "rutin setiap tahun.",

                        "Namun, di Tanah Toraja, upacara Rambu Solo' (kematian) "
                        "mulai mengalami perubahan bentuk. Beberapa keluarga "
                        "mulai menyederhanakan prosesi upacara karena alasan "
                        "ekonomi. Hal ini menunjukkan adanya kompromi antara "
                        "pelestarian tradisi dan tuntutan zaman.",

                        "Di Pontianak, komunitas Tionghoa masih mempertahankan "
                        "tradisi Imlek dan Cap Go Meh. Namun, generasi muda "
                        "komunitas Tionghoa mulai kehilangan kemampuan berbahasa "
                        "Hokkien dan Mandarin."
                    ]
                },
                {
                    "judul": "Faktor Pemengaruhi Pelestarian Budaya",
                    "sections": [
                        {
                            "judul": "Faktor Internal",
                            "isi": [
                                "Faktor internal yang mempengaruhi pelestarian "
                                "budaya meliputi:",
                                "1. Kesadaran masyarakat: tingkat kesadaran "
                                "masyarakat terhadap pentingnya melestarikan "
                                "budaya lokal masih bervariasi.",
                                "2. Ekonomi: kondisi ekonomi masyarakat berpengaruh "
                                "langsung terhadap kemampuan mereka dalam "
                                "melestarikan tradisi.",
                                "3. Pendidikan: tingkat pendidikan masyarakat "
                                "mempengaruhi cara pandang mereka terhadap "
                                "budaya lokal."
                            ]
                        },
                        {
                            "judul": "Faktor Eksternal",
                            "isi": [
                                "Faktor eksternal yang mempengaruhi pelestarian "
                                "budaya meliputi:",
                                "1. Kebijakan pemerintah: kebijakan pemerintah "
                                "daerah dalam mendukung pelestarian budaya.",
                                "2. Pengaruh globalisasi: masuknya budaya asing "
                                "melalui media digital.",
                                "3. Perkembangan teknologi: penggunaan teknologi "
                                "dalam pelestarian dan promosi budaya."
                            ]
                        }
                    ]
                },
                {
                    "judul": "Strategi Pengembangan Nilai Tradisional",
                    "isi": [
                        "Berdasarkan hasil analisis data, peneliti mengidentifikasi "
                        "beberapa strategi pengembangan nilai-nilai tradisional "
                        "yang efektif:",

                        "1. Pendidikan Budaya di Sekolah: Implementasi muatan "
                        "lokal dalam kurikulum sekolah terbukti efektif dalam "
                        "meningkatkan pengetahuan dan apresiasi siswa terhadap "
                        "budaya lokal.",

                        "2. Festival dan Event Budaya: Penyelenggaraan festival "
                        "budaya secara rutin menjadi sarana efektif untuk "
                        "memperkenalkan dan melestarikan budaya lokal.",

                        "3. Pemanfaatan Teknologi Digital: Penggunaan media "
                        "sosial dan platform digital untuk mempromosikan "
                        "budaya lokal kepada generasi muda.",

                        "4. Ekowisata Budaya: Pengembangan destinasi wisata "
                        "berbasis budaya yang memberikan manfaat ekonomi "
                        "bagi masyarakat lokal.",

                        "5. Pemberdayaan Komunitas: Pemberdayaan komunitas "
                        "lokal dalam pelestarian dan pengembangan budaya."
                    ]
                }
            ]
        },

        # BAB V: Penutup
        {
            "judul": "PENUTUP",
            "sections": [
                {
                    "judul": "Kesimpulan",
                    "isi": [
                        "Berdasarkan hasil penelitian dan pembahasan yang telah "
                        "diuraikan, dapat disimpulkan bahwa:",

                        "1. Keberagaman budaya Indonesia merupakan aset bangsa "
                        "yang sangat berharga. Indonesia memiliki lebih dari "
                        "300 suku bangsa, 700 bahasa daerah, dan beragam "
                        "agama yang hidup berdampingan secara damai.",

                        "2. Pelestarian keberagaman budaya menghadapi tantangan "
                        "serius berupa pengaruh globalisasi, urbanisasi, minimnya "
                        "regenerasi, dan komodifikasi budaya.",

                        "3. Strategi pengembangan nilai-nilai tradisional "
                        "memerlukan pendekatan holistik yang melibatkan "
                        "pendidikan, pemanfaatan teknologi, dan pemberdayaan "
                        "komunitas.",

                        "4. Peran masyarakat sangat penting dalam pelestarian "
                        "budaya. Kolaborasi antara pemerintah, masyarakat, "
                        "dan sektor swasta diperlukan untuk memastikan "
                        "keberlanjutan budaya lokal."
                    ]
                },
                {
                    "judul": "Saran",
                    "isi": [
                        "Berdasarkan temuan penelitian ini, penulis memberikan "
                        "saran sebagai berikut:",

                        "1. Bagi Pemerintah:",
                        "   a. Memperkuat kebijakan pelestarian budaya melalui "
                        "regulasi yang lebih komprehensif.",
                        "   b. Meningkatkan alokasi anggaran untuk program "
                        "pelestarian budaya.",
                        "   c. Mengembangkan museum digital untuk mendokumentasikan "
                        "warisan budaya.",

                        "2. Bagi Masyarakat:",
                        "   a. Meningkatkan kesadaran akan pentingnya melestarikan "
                        "budaya lokal.",
                        "   b. Aktif berpartisipasi dalam kegiatan pelestarian "
                        "budaya.",
                        "   c. Meneruskan tradisi dan adat istiadat kepada "
                        "generasi muda.",

                        "3. Bagi Akademisi:",
                        "   a. Melakukan penelitian lebih lanjut tentang "
                        "pelestarian budaya.",
                        "   b. Mengembangkan model pendidikan budaya yang "
                        "inovatif.",
                        "   c. Mempublikasikan hasil penelitian untuk "
                        "kontribusi ilmu pengetahuan.",

                        "4. Bagi Generasi Muda:",
                        "   a. Mempelajari dan menghargai budaya leluhur.",
                        "   b. Berpartisipasi aktif dalam kegiatan pelestarian "
                        "budaya.",
                        "   c. Menggunakan media digital untuk mempromosikan "
                        "budaya lokal."
                    ]
                }
            ]
        }
    ],

    # ============================================
    # DAFTAR PUSTAKA
    # ============================================
    daftar_pustaka=[
        {
            "penulis": "Koentjaraningrat",
            "judul": "Pengantar Ilmu Antropologi",
            "tahun": "2015",
            "edisi": "5",
            "penerbit": "Rineka Cipta"
        },
        {
            "penulis": "Miles, M. B., & Huberman, A. M.",
            "judul": "Qualitative Data Analysis: A Methods Sourcebook",
            "tahun": "2014",
            "edisi": "3",
            "penerbit": "Sage Publications"
        },
        {
            "penulis": "Geertz, C.",
            "judul": "The Interpretation of Cultures",
            "tahun": "2017",
            "edisi": "",
            "penerbit": "Basic Books"
        },
        {
            "penulis": "Badan Pusat Statistik",
            "judul": "Statistik Indonesia 2024",
            "tahun": "2024",
            "edisi": "",
            "penerbit": "BPS",
            "url": "https://www.bps.go.id"
        },
        {
            "penulis": "Kementerian Pendidikan dan Kebudayaan",
            "judul": "Laporan Kebudayaan Nasional 2024",
            "tahun": "2024",
            "edisi": "",
            "penerbit": "Kemdikbud"
        },
        {
            "penulis": "UNESCO",
            "judul": "World Culture Report 2023",
            "tahun": "2023",
            "edisi": "",
            "penerbit": "UNESCO Publishing"
        },
        {
            "penulis": "Moleong, L. J.",
            "judul": "Metodologi Penelitian Kualitatif",
            "tahun": "2018",
            "edisi": "36",
            "penerbit": "Remaja Rosdakarya"
        },
        {
            "penulis": "Sugiyono",
            "judul": "Metode Penelitian Kuantitatif, Kualitatif, dan R&D",
            "tahun": "2019",
            "edisi": "",
            "penerbit": "Alfabeta"
        },
        {
            "penulis": "Bachri, B. S.",
            "judul": "Pengantar Antropologi Indonesia",
            "tahun": "2020",
            "edisi": "2",
            "penerbit": "Rajawali Pers"
        },
        {
            "penulis": "Effendy, M.",
            "judul": "Ilmu Komunikasi: Suatu Pengantar",
            "tahun": "2022",
            "edisi": "4",
            "penerbit": "Remaja Rosdakarya"
        }
    ],

    # ============================================
    # LAMPIRAN
    # ============================================
    lampiran=[
        {
            "judul": "Lampiran 1: Kuesioner Penelitian",
            "isi": [
                "Berikut adalah panduan wawancara yang digunakan dalam penelitian ini:",
                "",
                "1. Bagaimana pendapat Saudara tentang kondisi kebudayaan di "
                "daerah Saudara saat ini?",
                "",
                "2. Apa saja tradisi atau adat istiadat yang masih dilestarikan "
                "di daerah Saudara?",
                "",
                "3. Bagaimana cara masyarakat melestarikan budaya lokal?",
                "",
                "4. Apa tantangan terbesar dalam pelestarian budaya?",
                "",
                "5. Bagaimana peran pemerintah daerah dalam pelestarian budaya?",
                "",
                "6. Apa pendapat Saudara tentang pengaruh budaya asing terhadap "
                "budaya lokal?",
                "",
                "7. Bagaimana cara menarik minat generasi muda terhadap budaya lokal?",
                "",
                "8. Apa saran Saudara untuk pelestarian budaya ke depan?"
            ]
        },
        {
            "judul": "Lampiran 2: Hasil Wawancara Ringkas",
            "isi": [
                "Ringkasan wawancara dengan subjek penelitian:",
                "",
                "Subjek 1 (Tokoh Adat Yogyakarta):",
                "\"Budaya Jawa masih kuat di pedesaan, tapi di kota mulai luntur. "
                "Anak-anak muda lebih suka budaya Korea daripada belajar tari Jawa.\"",
                "",
                "Subjek 2 (Budayawan Toraja):",
                "\"Upacara Rambu Solo' masih jalan, tapi sudah tidak sebesar "
                "dulu. Banyak yang menyederhanakan karena biaya mahal.\"",
                "",
                "Subjek 3 (Komunitas Tionghoa Pontianak):",
                "\"Imlek masih dirayakan, tapi generasi muda sudah tidak bisa "
                "bahasa Hokkien. Perlu ada sekolah bahasa untuk anak-anak.\""
            ]
        },
        {
            "judul": "Lampiran 3: Daftar Lokasi Penelitian",
            "isi": [
                "Lokasi penelitian yang dipilih:",
                "",
                "1. Yogyakarta, Daerah Istimewa Yogyakarta",
                "   - Keraton Yogyakarta",
                "   - Kampung Budaya Purbayan",
                "   - Desa Wisata Brayut",
                "",
                "2. Tana Toraja, Sulawesi Selatan",
                "   - Ke'te Kesu'",
                "   - Londa",
                "   - Palawa",
                "",
                "3. Pontianak, Kalimantan Barat",
                "   - Taman Budaya",
                "   - Kampung Kapuas",
                "   - Vihara Buddhiksa"
            ]
        }
    ]
)

# Simpan dokumen
output_path = "KTI_Keberagaman_Indonesia.docx"
doc.save(output_path)
print(f"KTI berhasil disimpan ke: {output_path}")
print(f"Jumlah bab: 5")
print(f"Judul: {doc.judul}")
print(f"Penulis: {doc.penulis}")
print(f"Institusi: {doc.institusi}")
