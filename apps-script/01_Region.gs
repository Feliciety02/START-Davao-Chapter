/* =====================================================================
 *  START SCHOLAR TRACKER  ·  01_Region
 *  ---------------------------------------------------------------------
 *  Inside: WHICH REGION: all Philippine regions, their provinces, ready-made school lists.
 *  To customise: Add your region's school list to REGION_SCHOOLS. Nothing else needs to change for a new region.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_01_Region() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 01_Region. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Region  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  REGION  ·  everything that depends on WHICH region the tracker is for
 *  ----------------------------------------------------------------
 *  Pick the region once: Settings tab › Region (drop-down), or answer the question when you click ①.
 *  Everything follows it by itself: form titles, emails, privacy notice, calendar name, reports,
 *  the province drop-down in the Schools tab and the ready-made school list (🛠 Setup tools).
 *
 *  TO ADD A REGION'S SCHOOL LIST: add a block to REGION_SCHOOLS below (same 9 columns as XI).
 *  Nothing else in the code needs to change.
 * ================================================================ */

// Philippine regions (PSA list, checked Oct 2026: Negros Island Region from 2024; Sulu moved to Region IX in 2025;
// Maguindanao split in 2022). provinces = the drop-down in Schools › Province (independent cities included).
var REGIONS = [
  { code: 'NCR', label: 'NCR', name: 'National Capital Region', also: ['metro manila', 'region ncr'],
    provinces: ['Caloocan', 'Las Piñas', 'Makati', 'Malabon', 'Mandaluyong', 'Manila', 'Marikina', 'Muntinlupa', 'Navotas',
                'Parañaque', 'Pasay', 'Pasig', 'Pateros', 'Quezon City', 'San Juan', 'Taguig', 'Valenzuela'] },
  { code: 'CAR', label: 'CAR', name: 'Cordillera Administrative Region', also: ['cordillera'],
    provinces: ['Abra', 'Apayao', 'Baguio City', 'Benguet', 'Ifugao', 'Kalinga', 'Mountain Province'] },
  { code: 'I', label: 'Region I', name: 'Ilocos Region', also: ['1', 'ilocos'],
    provinces: ['Ilocos Norte', 'Ilocos Sur', 'La Union', 'Pangasinan'] },
  { code: 'II', label: 'Region II', name: 'Cagayan Valley', also: ['2'],
    provinces: ['Batanes', 'Cagayan', 'Isabela', 'Nueva Vizcaya', 'Quirino'] },
  { code: 'III', label: 'Region III', name: 'Central Luzon', also: ['3'],
    provinces: ['Angeles City', 'Aurora', 'Bataan', 'Bulacan', 'Nueva Ecija', 'Olongapo City', 'Pampanga', 'Tarlac', 'Zambales'] },
  { code: 'IV-A', label: 'Region IV-A', name: 'CALABARZON', also: ['4a', 'iva', 'calabarzon'],
    provinces: ['Batangas', 'Cavite', 'Laguna', 'Lucena City', 'Quezon', 'Rizal'] },
  { code: 'MIMAROPA', label: 'MIMAROPA', name: 'Southwestern Tagalog Region', also: ['4b', 'ivb', 'region iv-b', 'mimaropa', 'southwestern tagalog'],
    provinces: ['Marinduque', 'Occidental Mindoro', 'Oriental Mindoro', 'Palawan', 'Puerto Princesa City', 'Romblon'] },
  { code: 'V', label: 'Region V', name: 'Bicol Region', also: ['5', 'bicol'],
    provinces: ['Albay', 'Camarines Norte', 'Camarines Sur', 'Catanduanes', 'Masbate', 'Naga City', 'Sorsogon'] },
  { code: 'VI', label: 'Region VI', name: 'Western Visayas', also: ['6'],
    provinces: ['Aklan', 'Antique', 'Capiz', 'Guimaras', 'Iloilo', 'Iloilo City'] },
  { code: 'NIR', label: 'NIR', name: 'Negros Island Region', also: ['18', 'xviii', 'region xviii', 'negros'],
    provinces: ['Bacolod City', 'Negros Occidental', 'Negros Oriental', 'Siquijor'] },
  { code: 'VII', label: 'Region VII', name: 'Central Visayas', also: ['7'],
    provinces: ['Bohol', 'Cebu', 'Cebu City', 'Lapu-Lapu City', 'Mandaue City'] },
  { code: 'VIII', label: 'Region VIII', name: 'Eastern Visayas', also: ['8'],
    provinces: ['Biliran', 'Eastern Samar', 'Leyte', 'Northern Samar', 'Ormoc City', 'Samar', 'Southern Leyte', 'Tacloban City'] },
  { code: 'IX', label: 'Region IX', name: 'Zamboanga Peninsula', also: ['9', 'zamboanga'],
    provinces: ['Isabela City', 'Sulu', 'Zamboanga City', 'Zamboanga del Norte', 'Zamboanga del Sur', 'Zamboanga Sibugay'] },
  { code: 'X', label: 'Region X', name: 'Northern Mindanao', also: ['10'],
    provinces: ['Bukidnon', 'Cagayan de Oro City', 'Camiguin', 'Iligan City', 'Lanao del Norte', 'Misamis Occidental', 'Misamis Oriental'] },
  { code: 'XI', label: 'Region XI', name: 'Davao Region', also: ['11', 'davao'],
    provinces: ['Davao City', 'Davao de Oro', 'Davao del Norte', 'Davao del Sur', 'Davao Occidental', 'Davao Oriental'] },
  { code: 'XII', label: 'Region XII', name: 'SOCCSKSARGEN', also: ['12', 'soccsksargen'],
    provinces: ['Cotabato', 'General Santos City', 'Sarangani', 'South Cotabato', 'Sultan Kudarat'] },
  { code: 'XIII', label: 'Region XIII', name: 'Caraga', also: ['13', 'caraga'],
    provinces: ['Agusan del Norte', 'Agusan del Sur', 'Butuan City', 'Dinagat Islands', 'Surigao del Norte', 'Surigao del Sur'] },
  { code: 'BARMM', label: 'BARMM', name: 'Bangsamoro Autonomous Region in Muslim Mindanao', also: ['bangsamoro', 'armm'],
    provinces: ['Basilan', 'Cotabato City', 'Lanao del Sur', 'Maguindanao del Norte', 'Maguindanao del Sur', 'Special Geographic Area', 'Tawi-Tawi'] },
];

// Ready-made school lists, by region code. Columns: School (acronym), Province, Official name, City/Municipality,
// Type, Website, Public phone, Public email, DOST-SEI basis.
// Region XI: Davao Region schools that fit DOST-SEI's enrolment rules: state universities/colleges,
// CHED Centers of Excellence/Development, or Level III-accredited / CHED-autonomous private HEIs.
// Public details from each school's website, CHED RO XI, CHED's COE/COD list and UMAP (checked Oct 2026).
// Blank phone/email = not published as plain text on the school's site.
var REGION_SCHOOLS = {
  XI: [
    ['USeP', 'Davao City', 'University of Southeastern Philippines', 'Davao City (Obrero campus)', 'SUC',
     'https://www.usep.edu.ph', '+63 82 227 8192', '', 'SUC; CHED COE/COD (Obrero, Tagum-Mabini)'],
    ['UP Mindanao', 'Davao City', 'University of the Philippines Mindanao', 'Mintal, Davao City', 'SUC (UP System)',
     'https://www.upmin.edu.ph', '(082) 293-0863 / 293-0185', '', 'SUC; CHED COD Biology'],
    ['DNSC', 'Davao del Norte', 'Davao del Norte State College', 'Panabo City', 'SUC',
     'https://dnsc.edu.ph', '(084) 645-3064', 'president@dnsc.edu.ph', 'SUC'],
    ['DOrSU', 'Davao Oriental', 'Davao Oriental State University', 'City of Mati', 'SUC',
     'https://www.dorsu.edu.ph', '+63 87 388 3195', '', 'SUC'],
    ['DSSC', 'Davao del Sur', 'Davao del Sur State College', 'Digos City', 'SUC',
     'https://dssc.edu.ph', '+63 976 122 2428', 'president@dssc.edu.ph', 'SUC'],
    ['DdOSC', 'Davao de Oro', 'Davao de Oro State College', 'Compostela', 'SUC',
     'https://ddosc.edu.ph', '+63 929 169 5328', 'records@ddosc.edu.ph', 'SUC'],
    ['SPAMAST', 'Davao Occidental', 'Southern Philippines Agri-Business and Marine and Aquatic School of Technology', 'Malita', 'SUC',
     'https://spamast.edu.ph', '(082) 272-3725', 'op@spamast.edu.ph', 'SUC'],
    ['AdDU', 'Davao City', 'Ateneo de Davao University', 'Davao City', 'Private',
     'https://www.addu.edu.ph', '(082) 221-2411', 'contactus@addu.edu.ph', 'CHED COD IT; CHED autonomous'],
    ['UM', 'Davao City', 'The University of Mindanao', 'Davao City (Bolton St.)', 'Private',
     'https://umindanao.edu.ph', '(082) 227-5456', '', 'CHED COE/COD (IT, CpE, EE, ME); CHED autonomous'],
    ['UM Digos', 'Davao del Sur', 'UM Digos College', 'Digos City', 'Private',
     'https://digos.umindanao.edu.ph', '(082) 553-2914', '', 'CHED COD IT; CHED autonomous'],
    ['UM Tagum', 'Davao del Norte', 'UM Tagum College', 'Tagum City', 'Private',
     'https://tagum.umindanao.edu.ph', '', '', 'CHED COD Teacher Ed; CHED autonomous'],
    ['UIC', 'Davao City', 'University of the Immaculate Conception', 'Davao City (Fr. Selga St.)', 'Private',
     'https://www.uic.edu.ph', '(082) 221-8090', '', 'CHED COD IT; CHED autonomous'],
    ['CJC', 'Davao del Sur', 'Cor Jesu College', 'Digos City', 'Private',
     'https://cjc.edu.ph', '(082) 553-2433', 'customerservice@cjc.edu.ph', 'CHED autonomous (confirm with DOST-SEI)'],
    ],
};
var REGION_XI_SCHOOLS = REGION_SCHOOLS.XI;     // older name, kept so nothing breaks
/* ---------- every Philippine province and its cities / municipalities (for the sign-up form's hometown) ----------
 * Source: Philippine Standard Geographic Code (PSA), 2025 2nd quarter, via the open @jobuntux/psgc dataset (MIT).
 * Highly urbanized cities are listed under their province (Davao City under Davao del Sur); Metro Manila is one entry.
 * 82 provinces + Metro Manila + the BARMM Special Geographic Area, 1,656 cities and municipalities. Checked: 2026-10-23. */
var PH_PLACES = {"Abra":["Bangued","Boliney","Bucay","Bucloc","Daguioman","Danglas","Dolores","La Paz","Lacub","Lagangilang","Lagayan","Langiden","Licuan-Baay","Luba","Malibcong","Manabo","Peñarrubia","Pidigan","Pilar","Sallapadan","San Isidro","San Juan","San Quintin","Tayum","Tineg","Tubo","Villaviciosa"],"Agusan del Norte":["Buenavista","Butuan City","Cabadbaran City","Carmen","Jabonga","Kitcharao","Las Nieves","Magallanes","Nasipit","Remedios T. Romualdez","Santiago","Tubay"],"Agusan del Sur":["Bayugan City","Bunawan","Esperanza","La Paz","Loreto","Prosperidad","Rosario","San Francisco","San Luis","Santa Josefa","Sibagat","Talacogon","Trento","Veruela"],"Aklan":["Altavas","Balete","Banga","Batan","Buruanga","Ibajay","Kalibo","Lezo","Libacao","Madalag","Makato","Malay","Malinao","Nabas","New Washington","Numancia","Tangalan"],"Albay":["Bacacay","Camalig","Daraga","Guinobatan","Jovellar","Legazpi City","Libon","Ligao City","Malilipot","Malinao","Manito","Oas","Pio Duran","Polangui","Rapu-Rapu","Santo Domingo","Tabaco City","Tiwi"],"Antique":["Anini-Y","Barbaza","Belison","Bugasong","Caluya","Culasi","Hamtic","Laua-An","Libertad","Pandan","Patnongon","San Jose","San Remigio","Sebaste","Sibalom","Tibiao","Tobias Fornier","Valderrama"],"Apayao":["Calanasan","Conner","Flora","Kabugao","Luna","Pudtol","Santa Marcela"],"Aurora":["Baler","Casiguran","Dilasag","Dinalungan","Dingalan","Dipaculao","Maria Aurora","San Luis"],"Basilan":["Akbar","Al-Barka","Hadji Mohammad Ajul","Hadji Muhtamad","Isabela City","Lamitan City","Lantawan","Maluso","Sumisip","Tabuan-Lasa","Tipo-Tipo","Tuburan","Ungkaya Pukan"],"Bataan":["Abucay","Bagac","Balanga City","Dinalupihan","Hermosa","Limay","Mariveles","Morong","Orani","Orion","Pilar","Samal"],"Batanes":["Basco","Itbayat","Ivana","Mahatao","Sabtang","Uyugan"],"Batangas":["Agoncillo","Alitagtag","Balayan","Balete","Batangas City","Bauan","Calaca City","Calatagan","Cuenca","Ibaan","Laurel","Lemery","Lian","Lipa City","Lobo","Mabini","Malvar","Mataasnakahoy","Nasugbu","Padre Garcia","Rosario","San Jose","San Juan","San Luis","San Nicolas","San Pascual","Santa Teresita","Sto. Tomas City","Taal","Talisay","Tanauan City","Taysan","Tingloy","Tuy"],"Benguet":["Atok","Baguio City","Bakun","Bokod","Buguias","Itogon","Kabayan","Kapangan","Kibungan","La Trinidad","Mankayan","Sablan","Tuba","Tublay"],"Biliran":["Almeria","Biliran","Cabucgayan","Caibiran","Culaba","Kawayan","Maripipi","Naval"],"Bohol":["Alburquerque","Alicia","Anda","Antequera","Baclayon","Balilihan","Batuan","Bien Unido","Bilar","Buenavista","Calape","Candijay","Carmen","Catigbian","Clarin","Corella","Cortes","Dagohoy","Danao","Dauis","Dimiao","Duero","Garcia Hernandez","Getafe","Guindulman","Inabanga","Jagna","Lila","Loay","Loboc","Loon","Mabini","Maribojoc","Panglao","Pilar","President Carlos P. Garcia","Sagbayan","San Isidro","San Miguel","Sevilla","Sierra Bullones","Sikatuna","Tagbilaran City","Talibon","Trinidad","Tubigon","Ubay","Valencia"],"Bukidnon":["Baungon","Cabanglasan","Damulog","Dangcagan","Don Carlos","Impasug-ong","Kadingilan","Kalilangan","Kibawe","Kitaotao","Lantapan","Libona","Malaybalay City","Malitbog","Manolo Fortich","Maramag","Pangantucan","Quezon","San Fernando","Sumilao","Talakag","Valencia City"],"Bulacan":["Angat","Balagtas","Baliwag City","Bocaue","Bulacan","Bustos","Calumpit","Doña Remedios Trinidad","Guiguinto","Hagonoy","Malolos City","Marilao","Meycauayan City","Norzagaray","Obando","Pandi","Paombong","Plaridel","Pulilan","San Ildefonso","San Jose Del Monte City","San Miguel","San Rafael","Santa Maria"],"Cagayan":["Abulug","Alcala","Allacapan","Amulung","Aparri","Baggao","Ballesteros","Buguey","Calayan","Camalaniugan","Claveria","Enrile","Gattaran","Gonzaga","Iguig","Lal-Lo","Lasam","Pamplona","Peñablanca","Piat","Rizal","Sanchez-Mira","Santa Ana","Santa Praxedes","Santa Teresita","Santo Niño","Solana","Tuao","Tuguegarao City"],"Camarines Norte":["Basud","Capalonga","Daet","Jose Panganiban","Labo","Mercedes","Paracale","San Lorenzo Ruiz","San Vicente","Santa Elena","Talisay","Vinzons"],"Camarines Sur":["Baao","Balatan","Bato","Bombon","Buhi","Bula","Cabusao","Calabanga","Camaligan","Canaman","Caramoan","Del Gallego","Gainza","Garchitorena","Goa","Iriga City","Lagonoy","Libmanan","Lupi","Magarao","Milaor","Minalabac","Nabua","Naga City","Ocampo","Pamplona","Pasacao","Pili","Presentacion","Ragay","Sagñay","San Fernando","San Jose","Sipocot","Siruma","Tigaon","Tinambac"],"Camiguin":["Catarman","Guinsiliban","Mahinog","Mambajao","Sagay"],"Capiz":["Cuartero","Dao","Dumalag","Dumarao","Ivisan","Jamindan","Ma-Ayon","Mambusao","Panay","Panitan","Pilar","Pontevedra","President Roxas","Roxas City","Sapi-An","Sigma","Tapaz"],"Catanduanes":["Bagamanoc","Baras","Bato","Caramoran","Gigmoto","Pandan","Panganiban","San Andres","San Miguel","Viga","Virac"],"Cavite":["Alfonso","Amadeo","Bacoor City","Carmona City","Cavite City","Dasmariñas City","Gen. Mariano Alvarez","General Emilio Aguinaldo","General Trias City","Imus City","Indang","Kawit","Magallanes","Maragondon","Mendez","Naic","Noveleta","Rosario","Silang","Tagaytay City","Tanza","Ternate","Trece Martires City"],"Cebu":["Alcantara","Alcoy","Alegria","Aloguinsan","Argao","Asturias","Badian","Balamban","Bantayan","Barili","Bogo City","Boljoon","Borbon","Carcar City","Carmen","Catmon","Cebu City","Compostela","Consolacion","Cordova","Daanbantayan","Dalaguete","Danao City","Dumanjug","Ginatilan","Lapu-Lapu City","Liloan","Madridejos","Malabuyoc","Mandaue City","Medellin","Minglanilla","Moalboal","Naga City","Oslob","Pilar","Pinamungajan","Poro","Ronda","Samboan","San Fernando","San Francisco","San Remigio","Santa Fe","Santander","Sibonga","Sogod","Tabogon","Tabuelan","Talisay City","Toledo City","Tuburan","Tudela"],"Cotabato":["Alamada","Aleosan","Antipas","Arakan","Banisilan","Carmen","Kabacan","Kidapawan City","Libungan","M'Lang","Magpet","Makilala","Matalam","Midsayap","Pigkawayan","Pikit","President Roxas","Tulunan"],"Davao Occidental":["Don Marcelino","Jose Abad Santos","Malita","Santa Maria","Sarangani"],"Davao Oriental":["Baganga","Banaybanay","Boston","Caraga","Cateel","Governor Generoso","Lupon","Manay","Mati City","San Isidro","Tarragona"],"Davao de Oro":["Compostela","Laak","Mabini","Maco","Maragusan","Mawab","Monkayo","Montevista","Nabunturan","New Bataan","Pantukan"],"Davao del Norte":["Asuncion","Braulio E. Dujali","Carmen","Island Garden City of Samal","Kapalong","New Corella","Panabo City","San Isidro","Santo Tomas","Tagum City","Talaingod"],"Davao del Sur":["Bansalan","Davao City","Digos City","Hagonoy","Kiblawan","Magsaysay","Malalag","Matanao","Padada","Santa Cruz","Sulop"],"Dinagat Islands":["Basilisa","Cagdianao","Dinagat","Libjo","Loreto","San Jose","Tubajon"],"Eastern Samar":["Arteche","Balangiga","Balangkayan","Borongan City","Can-Avid","Dolores","General Macarthur","Giporlos","Guiuan","Hernani","Jipapad","Lawaan","Llorente","Maslog","Maydolong","Mercedes","Oras","Quinapondan","Salcedo","San Julian","San Policarpo","Sulat","Taft"],"Guimaras":["Buenavista","Jordan","Nueva Valencia","San Lorenzo","Sibunag"],"Ifugao":["Aguinaldo","Alfonso Lista","Asipulo","Banaue","Hingyon","Hungduan","Kiangan","Lagawe","Lamut","Mayoyao","Tinoc"],"Ilocos Norte":["Adams","Bacarra","Badoc","Bangui","Banna","Batac City","Burgos","Carasi","Currimao","Dingras","Dumalneg","Laoag City","Marcos","Nueva Era","Pagudpud","Paoay","Pasuquin","Piddig","Pinili","San Nicolas","Sarrat","Solsona","Vintar"],"Ilocos Sur":["Alilem","Banayoyo","Bantay","Burgos","Cabugao","Candon City","Caoayan","Cervantes","Galimuyod","Gregorio del Pilar","Lidlidda","Magsingal","Nagbukel","Narvacan","Quirino","Salcedo","San Emilio","San Esteban","San Ildefonso","San Juan","San Vicente","Santa","Santa Catalina","Santa Cruz","Santa Lucia","Santa Maria","Santiago","Santo Domingo","Sigay","Sinait","Sugpon","Suyo","Tagudin","Vigan City"],"Iloilo":["Ajuy","Alimodian","Anilao","Badiangan","Balasan","Banate","Barotac Nuevo","Barotac Viejo","Batad","Bingawan","Cabatuan","Calinog","Carles","Concepcion","Dingle","Dueñas","Dumangas","Estancia","Guimbal","Igbaras","Iloilo City","Janiuay","Lambunao","Leganes","Lemery","Leon","Maasin","Miagao","Mina","New Lucena","Oton","Passi City","Pavia","Pototan","San Dionisio","San Enrique","San Joaquin","San Miguel","San Rafael","Santa Barbara","Sara","Tigbauan","Tubungan","Zarraga"],"Isabela":["Alicia","Angadanan","Aurora","Benito Soliven","Burgos","Cabagan","Cabatuan","Cauayan City","Cordon","Delfin Albano","Dinapigue","Divilacan","Echague","Gamu","Ilagan City","Jones","Luna","Maconacon","Mallig","Naguilian","Palanan","Quezon","Quirino","Ramon","Reina Mercedes","Roxas","San Agustin","San Guillermo","San Isidro","San Manuel","San Mariano","San Mateo","San Pablo","Santa Maria","Santiago City","Santo Tomas","Tumauini"],"Kalinga":["Balbalan","Lubuagan","Pasil","Pinukpuk","Rizal","Tabuk City","Tanudan","Tinglayan"],"La Union":["Agoo","Aringay","Bacnotan","Bagulin","Balaoan","Bangar","Bauang","Burgos","Caba","Luna","Naguilian","Pugo","Rosario","San Fernando City","San Gabriel","San Juan","Santo Tomas","Santol","Sudipen","Tubao"],"Laguna":["Alaminos","Bay","Biñan City","Cabuyao City","Calamba City","Calauan","Cavinti","Famy","Kalayaan","Liliw","Los Baños","Luisiana","Lumban","Mabitac","Magdalena","Majayjay","Nagcarlan","Paete","Pagsanjan","Pakil","Pangil","Pila","Rizal","San Pablo City","San Pedro City","Santa Cruz","Santa Maria","Santa Rosa City","Siniloan","Victoria"],"Lanao del Norte":["Bacolod","Baloi","Baroy","Iligan City","Kapatagan","Kauswagan","Kolambugan","Lala","Linamon","Magsaysay","Maigo","Matungao","Munai","Nunungan","Pantao Ragat","Pantar","Poona Piagapo","Salvador","Sapad","Sultan Naga Dimaporo","Tagoloan","Tangcal","Tubod"],"Lanao del Sur":["Amai Manabilang","Bacolod-Kalawi","Balabagan","Balindong","Bayang","Binidayan","Buadiposo-Buntong","Bubong","Butig","Calanogas","Ditsaan-Ramain","Ganassi","Kapai","Kapatagan","Lumba-Bayabao","Lumbaca-Unayan","Lumbatan","Lumbayanague","Madalum","Madamba","Maguing","Malabang","Marantao","Marawi City","Marogong","Masiu","Mulondo","Pagayawan","Piagapo","Picong","Poona Bayabao","Pualas","Saguiaran","Sultan Dumalondong","Tagoloan Ii","Tamparan","Taraka","Tubaran","Tugaya","Wao"],"Leyte":["Abuyog","Alangalang","Albuera","Babatngon","Barugo","Bato","Baybay City","Burauen","Calubian","Capoocan","Carigara","Dagami","Dulag","Hilongos","Hindang","Inopacan","Isabel","Jaro","Javier","Julita","Kananga","La Paz","Leyte","Macarthur","Mahaplag","Matag-Ob","Matalom","Mayorga","Merida","Ormoc City","Palo","Palompon","Pastrana","San Isidro","San Miguel","Santa Fe","Tabango","Tabontabon","Tacloban City","Tanauan","Tolosa","Tunga","Villaba"],"Maguindanao del Norte":["Barira","Buldon","Cotabato City","Datu Blah T. Sinsuat","Datu Odin Sinsuat","Kabuntalan","Matanog","Northern Kabuntalan","Parang","Sultan Kudarat","Sultan Mastura","Talitay","Upi"],"Maguindanao del Sur":["Ampatuan","Buluan","Datu Abdullah Sangki","Datu Anggal Midtimbang","Datu Hoffer Ampatuan","Datu Paglas","Datu Piang","Datu Salibo","Datu Saudi Ampatuan","Datu Unsay","Gen. S.K. Pendatun","Guindulungan","Mamasapano","Mangudadatu","Pagagawan","Pagalungan","Paglat","Pandag","Rajah Buayan","Shariff Aguak","Shariff Saydona Mustapha","South Upi","Sultan Sa Barongis","Talayan"],"Marinduque":["Boac","Buenavista","Gasan","Mogpog","Santa Cruz","Torrijos"],"Masbate":["Aroroy","Baleno","Balud","Batuan","Cataingan","Cawayan","Claveria","Dimasalang","Esperanza","Mandaon","Masbate City","Milagros","Mobo","Monreal","Palanas","Pio V. Corpus","Placer","San Fernando","San Jacinto","San Pascual","Uson"],"Metro Manila":["Binondo","Caloocan City","Ermita","Intramuros","Las Piñas City","Makati City","Malabon City","Malate","Mandaluyong City","Manila City","Marikina City","Muntinlupa City","Navotas City","Paco","Pandacan","Parañaque City","Pasay City","Pasig City","Pateros","Port Area","Quezon City","Quiapo","Sampaloc","San Juan City","San Miguel","San Nicolas","Santa Ana","Santa Cruz","Taguig City","Tondo I/II","Valenzuela City"],"Misamis Occidental":["Aloran","Baliangao","Bonifacio","Calamba","Clarin","Concepcion","Don Victoriano Chiongbian","Jimenez","Lopez Jaena","Oroquieta City","Ozamiz City","Panaon","Plaridel","Sapang Dalaga","Sinacaban","Tangub City","Tudela"],"Misamis Oriental":["Alubijid","Balingasag","Balingoan","Binuangan","Cagayan de Oro City","Claveria","El Salvador City","Gingoog City","Gitagum","Initao","Jasaan","Kinoguitan","Lagonglong","Laguindingan","Libertad","Lugait","Magsaysay","Manticao","Medina","Naawan","Opol","Salay","Sugbongcogon","Tagoloan","Talisayan","Villanueva"],"Mountain Province":["Barlig","Bauko","Besao","Bontoc","Natonin","Paracelis","Sabangan","Sadanga","Sagada","Tadian"],"Negros Occidental":["Bacolod City","Bago City","Binalbagan","Cadiz City","Calatrava","Candoni","Cauayan","Enrique B. Magalona","Escalante City","Himamaylan City","Hinigaran","Hinoba-an","Ilog","Isabela","Kabankalan City","La Carlota City","La Castellana","Manapla","Moises Padilla","Murcia","Pontevedra","Pulupandan","Sagay City","Salvador Benedicto","San Carlos City","San Enrique","Silay City","Sipalay City","Talisay City","Toboso","Valladolid","Victorias City"],"Negros Oriental":["Amlan","Ayungon","Bacong","Bais City","Basay","Bayawan City","Bindoy","Canlaon City","Dauin","Dumaguete City","Guihulngan City","Jimalalud","La Libertad","Mabinay","Manjuyod","Pamplona","San Jose","Santa Catalina","Siaton","Sibulan","Tanjay City","Tayasan","Valencia","Vallehermoso","Zamboanguita"],"Northern Samar":["Allen","Biri","Bobon","Capul","Catarman","Catubig","Gamay","Laoang","Lapinig","Las Navas","Lavezares","Lope De Vega","Mapanas","Mondragon","Palapag","Pambujan","Rosario","San Antonio","San Isidro","San Jose","San Roque","San Vicente","Silvino Lobos","Victoria"],"Nueva Ecija":["Aliaga","Bongabon","Cabanatuan City","Cabiao","Carranglan","Cuyapo","Gabaldon","Gapan City","General Mamerto Natividad","General Tinio","Guimba","Jaen","Laur","Licab","Llanera","Lupao","Nampicuan","Palayan City","Pantabangan","Peñaranda","Quezon","Rizal","San Antonio","San Isidro","San Jose City","San Leonardo","Santa Rosa","Santo Domingo","Science City of Muñoz","Talavera","Talugtug","Zaragoza"],"Nueva Vizcaya":["Alfonso Castaneda","Ambaguio","Aritao","Bagabag","Bambang","Bayombong","Diadi","Dupax del Norte","Dupax del Sur","Kasibu","Kayapa","Quezon","Santa Fe","Solano","Villaverde"],"Occidental Mindoro":["Abra De Ilog","Calintaan","Looc","Lubang","Magsaysay","Mamburao","Paluan","Rizal","Sablayan","San Jose","Santa Cruz"],"Oriental Mindoro":["Baco","Bansud","Bongabong","Bulalacao","Calapan City","Gloria","Mansalay","Naujan","Pinamalayan","Pola","Puerto Galera","Roxas","San Teodoro","Socorro","Victoria"],"Palawan":["Aborlan","Agutaya","Araceli","Balabac","Bataraza","Brooke's Point","Busuanga","Cagayancillo","Coron","Culion","Cuyo","Dr. Jose P. Rizal","Dumaran","El Nido","Kalayaan","Linapacan","Magsaysay","Narra","Puerto Princesa City","Quezon","Roxas","San Vicente","Sofronio Española","Taytay"],"Pampanga":["Angeles City","Apalit","Arayat","Bacolor","Candaba","Floridablanca","Guagua","Lubao","Mabalacat City","Macabebe","Magalang","Masantol","Mexico","Minalin","Porac","San Fernando City","San Luis","San Simon","Santa Ana","Santa Rita","Sasmuan","Sto. Tomas"],"Pangasinan":["Agno","Aguilar","Alaminos City","Alcala","Anda","Asingan","Balungao","Bani","Basista","Bautista","Bayambang","Binalonan","Binmaley","Bolinao","Bugallon","Burgos","Calasiao","Dagupan City","Dasol","Infanta","Labrador","Laoac","Lingayen","Mabini","Malasiqui","Manaoag","Mangaldan","Mangatarem","Mapandan","Natividad","Pozorrubio","Rosales","San Carlos City","San Fabian","San Jacinto","San Manuel","San Nicolas","San Quintin","Santa Barbara","Santa Maria","Santo Tomas","Sison","Sual","Tayug","Umingan","Urbiztondo","Urdaneta City","Villasis"],"Quezon":["Agdangan","Alabat","Atimonan","Buenavista","Burdeos","Calauag","Candelaria","Catanauan","Dolores","General Luna","General Nakar","Guinayangan","Gumaca","Infanta","Jomalig","Lopez","Lucban","Lucena City","Macalelon","Mauban","Mulanay","Padre Burgos","Pagbilao","Panukulan","Patnanungan","Perez","Pitogo","Plaridel","Polillo","Quezon","Real","Sampaloc","San Andres","San Antonio","San Francisco","San Narciso","Sariaya","Tagkawayan","Tayabas City","Tiaong","Unisan"],"Quirino":["Aglipay","Cabarroguis","Diffun","Maddela","Nagtipunan","Saguday"],"Rizal":["Angono","Antipolo City","Baras","Binangonan","Cainta","Cardona","Jala-Jala","Morong","Pililla","Rodriguez","San Mateo","Tanay","Taytay","Teresa"],"Romblon":["Alcantara","Banton","Cajidiocan","Calatrava","Concepcion","Corcuera","Ferrol","Looc","Magdiwang","Odiongan","Romblon","San Agustin","San Andres","San Fernando","San Jose","Santa Fe","Santa Maria"],"Samar":["Almagro","Basey","Calbayog City","Calbiga","Catbalogan City","Daram","Gandara","Hinabangan","Jiabong","Marabut","Matuguinao","Motiong","Pagsanghan","Paranas","Pinabacdao","San Jorge","San Jose De Buan","San Sebastian","Santa Margarita","Santa Rita","Santo Niño","Tagapul-An","Talalora","Tarangnan","Villareal","Zumarraga"],"Sarangani":["Alabel","Glan","Kiamba","Maasim","Maitum","Malapatan","Malungon"],"Siquijor":["Enrique Villanueva","Larena","Lazi","Maria","San Juan","Siquijor"],"Sorsogon":["Barcelona","Bulan","Bulusan","Casiguran","Castilla","Donsol","Gubat","Irosin","Juban","Magallanes","Matnog","Pilar","Prieto Diaz","Santa Magdalena","Sorsogon City"],"South Cotabato":["Banga","General Santos City","Koronadal City","Lake Sebu","Norala","Polomolok","Santo Niño","Surallah","T'Boli","Tampakan","Tantangan","Tupi"],"Southern Leyte":["Anahawan","Bontoc","Hinunangan","Hinundayan","Libagon","Liloan","Limasawa","Maasin City","Macrohon","Malitbog","Padre Burgos","Pintuyan","Saint Bernard","San Francisco","San Juan","San Ricardo","Silago","Sogod","Tomas Oppus"],"Special Geographic Area (BARMM)":["Kadayangan","Kapalawan","Ligawasan","Malidegao","Nabalawag","Old Kaabakan","Pahamuddin","Tugunan"],"Sultan Kudarat":["Bagumbayan","Columbio","Esperanza","Isulan","Kalamansig","Lambayong","Lebak","Lutayan","Palimbang","President Quirino","Sen. Ninoy Aquino","Tacurong City"],"Sulu":["Hadji Panglima Tahil","Indanan","Jolo","Kalingalan Caluang","Lugus","Luuk","Maimbung","Old Panamao","Omar","Pandami","Panglima Estino","Pangutaran","Parang","Pata","Patikul","Siasi","Talipao","Tapul","Tongkil"],"Surigao del Norte":["Alegria","Bacuag","Burgos","Claver","Dapa","Del Carmen","General Luna","Gigaquit","Mainit","Malimono","Pilar","Placer","San Benito","San Francisco","San Isidro","Santa Monica","Sison","Socorro","Surigao City","Tagana-An","Tubod"],"Surigao del Sur":["Barobo","Bayabas","Bislig City","Cagwait","Cantilan","Carmen","Carrascal","Cortes","Hinatuan","Lanuza","Lianga","Lingig","Madrid","Marihatag","San Agustin","San Miguel","Tagbina","Tago","Tandag City"],"Tarlac":["Anao","Bamban","Camiling","Capas","Concepcion","Gerona","La Paz","Mayantoc","Moncada","Paniqui","Pura","Ramos","San Clemente","San Jose","San Manuel","Santa Ignacia","Tarlac City","Victoria"],"Tawi-Tawi":["Bongao","Languyan","Mapun","Panglima Sugala","Sapa-Sapa","Sibutu","Simunul","Sitangkai","South Ubian","Tandubas","Turtle Islands"],"Zambales":["Botolan","Cabangan","Candelaria","Castillejos","Iba","Masinloc","Olongapo City","Palauig","San Antonio","San Felipe","San Marcelino","San Narciso","Santa Cruz","Subic"],"Zamboanga Sibugay":["Alicia","Buug","Diplahan","Imelda","Ipil","Kabasalan","Mabuhay","Malangas","Naga","Olutanga","Payao","Roseller Lim","Siay","Talusan","Titay","Tungawan"],"Zamboanga del Norte":["Baliguian","Dapitan City","Dipolog City","Godod","Gutalac","Jose Dalman","Kalawit","Katipunan","La Libertad","Labason","Leon T. Postigo","Liloy","Manukan","Mutia","Piñan","Polanco","Pres. Manuel A. Roxas","Rizal","Salug","Sergio Osmeña Sr.","Siayan","Sibuco","Sibutad","Sindangan","Siocon","Sirawai","Tampilisan"],"Zamboanga del Sur":["Aurora","Bayog","Dimataling","Dinas","Dumalinao","Dumingag","Guipos","Josefina","Kumalarang","Labangan","Lakewood","Lapuyan","Mahayag","Margosatubig","Midsalip","Molave","Pagadian City","Pitogo","Ramon Magsaysay","San Miguel","San Pablo","Sominot","Tabina","Tambulig","Tigbao","Tukuran","Vincenzo A. Sagun","Zamboanga City"]};

// What the region drop-down shows, e.g. "Region XI – Davao Region"
function regionChoice_(r) { return r.label === r.name ? r.label : r.label + ' – ' + r.name; }
function regionChoices_() { return REGIONS.map(regionChoice_); }

// Finds the region from whatever was typed: "Region XI – Davao Region", "Region 11", "XI", "davao", "NCR"…
function findRegion_(text) {
  const norm = x => String(x || '').toLowerCase().replace(/ñ/g, 'n').replace(/[^a-z0-9]/g, '');
  const t = norm(text);
  if (!t) return null;
  const arabic = { I: 1, II: 2, III: 3, 'IV-A': '4a', V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10, XI: 11, XII: 12, XIII: 13 };
  return REGIONS.find(r => {
    const keys = [regionChoice_(r), r.label, r.name, r.code, 'region ' + r.code].concat(r.also || [])
      .concat(arabic[r.code] !== undefined ? ['region ' + arabic[r.code]] : []);
    return keys.some(k => norm(k) === t);
  }) || null;
}

// The tracker's region (from the Settings tab). Unknown / blank: label = what was typed, no lists.
function regionProfile_(st) {
  let typed = '';
  try { st = st || getSettings_(); typed = st.regionRaw !== undefined ? st.regionRaw : st.region; } catch (e) { typed = ''; }
  const r = findRegion_(typed);
  if (r) return { code: r.code, label: r.label, name: r.name, provinces: r.provinces, schools: REGION_SCHOOLS[r.code] || [] };
  const t = String(typed || '').trim();
  return { code: '', label: t, name: t, provinces: [], schools: [] };
}

// One setting straight from the Settings tab (no defaults)
function readSetting_(key) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TAB.settings);
    const i = SETTINGS.findIndex(s => s[0] === key);
    return sh && i !== -1 ? sh.getRange(i + 2, 2).getValue() : '';
  } catch (e) { return ''; }
}

// ① asks once when no region is chosen yet (a new tracker, or the blank copy for another RR)
function askRegionIfMissing_() {
  if (findRegion_(readSetting_('region'))) return true;
  let ui; try { ui = SpreadsheetApp.getUi(); if (!ui.prompt) return false; } catch (e) { return false; }
  const list = REGIONS.map((r, i) => `${i + 1}.  ${regionChoice_(r)}`).join('\n');
  const res = ui.prompt('Which region is this tracker for?', 'Type the NUMBER of your region:\n\n' + list +
                        '\n\n(You can change it later: Settings tab › Region.)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return false;
  const txt = res.getResponseText().trim();
  const r = REGIONS[Number(txt) - 1] || findRegion_(txt);
  if (!r) { ui.alert('START Tracker', 'That is not on the list. Pick your region later in the Settings tab › Region.', ui.ButtonSet.OK); return false; }
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.settings);
  if (sh) sh.getRange(settingCell_('region').split('!')[1]).setValue(regionChoice_(r));
  return true;
}

// Makes the sheet follow the region: Settings drop-down + Schools › Province drop-down
function applyRegion_(ss) {
  const set = ss.getSheetByName(TAB.settings);
  if (set) {
    const cell = set.getRange(settingCell_('region').split('!')[1]);
    cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(regionChoices_(), true).setAllowInvalid(true)
      .setHelpText('Pick your region. Everything in the tracker follows it.').build());
    const r = findRegion_(cell.getValue());
    if (r && cell.getValue() !== regionChoice_(r)) cell.setValue(regionChoice_(r));      // "Region XI" -> "Region XI – Davao Region"
  }
  const sh = ss.getSheetByName(TAB.schools), r = regionProfile_();
  if (sh) {
    const col = sh.getRange(2, 2, CONFIG.maxSchools, 1);
    if (r.provinces.length) col.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(r.provinces, true)
      .setAllowInvalid(true).setHelpText(`Province (${r.label}). Pick from the list, or type another.`).build());
    else col.clearDataValidations();
  }
  // The file name follows the region ("START Scholar Tracker – Region VII"), so every region's copy is easy to tell apart.
  // Only START names (or a fresh copy / the blank template) are renamed: a name you chose yourself is kept.
  try {
    const now = ss.getName(), want = 'START Scholar Tracker' + (r.label ? ' – ' + r.label : '');
    if (r.label && now !== want && /^(Copy of )?START[ _-]*(Scholar[ _-]*)?Tracker/i.test(now)) ss.rename(want);
  } catch (e) { Logger.log('File name: ' + e.message); }
  docProps_().setProperty('REGION_APPLIED', r.code + '|' + r.label);
  return r;
}

var SCHOOL_INFO_HEADERS = ['Official name', 'City / Municipality', 'Type', 'Website',
                             'Public phone', 'Public email', 'DOST-SEI basis'];

function loadSchoolList() {
  requireSetup_();
  const r = regionProfile_();
  if (!r.schools.length) {
    notify_((r.code ? `No ready-made school list for ${r.label} yet` : 'Choose your region first (Settings › Region)') +
            '\n\nType your schools in the Schools tab: short name in A, province in B.');
    return;
  }
  const n = writeSchoolList_(SpreadsheetApp.getActive());
  try { updateDropdown_(); } catch (e) { Logger.log('Dropdown not refreshed: ' + e.message); }
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(TAB.schools));
  if (!n) return;
  notify_(`✅ Added ${n} ${regionProfile_().name} schools\n\nWhat you typed before was kept. Check it against ONSSE's masterlist.`);
}

function writeSchoolList_(ss) {
  const sh = ss.getSheetByName(TAB.schools);
  ensureInfoHeaders_(sh);
  const names = sh.getRange(2, 1, CONFIG.maxSchools, 1).getValues().map(r => String(r[0]).trim());
  const list = regionProfile_().schools;
  list.forEach(s => {
    let i = names.findIndex(x => x.toLowerCase() === s[0].toLowerCase());
    if (i === -1) i = names.findIndex(x => !x);
    if (i === -1) throw new Error('The Schools tab is full (' + CONFIG.maxSchools + ' schools).');
    names[i] = s[0];
    sh.getRange(i + 2, 1, 1, 2).setValues([[s[0], s[1]]]);
    sh.getRange(i + 2, 9, 1, 7).setValues([s.slice(2)]);
  });
  return list.length;
}

var LINK_COL = 16;   // P: each school's own form link, school already selected

// Makes sure the Schools headings exist before the school list or the form sync write into columns I–P (headings come from columnSpecs_)
function ensureInfoHeaders_(sh) {
  if (isMarked_(sh.getRange(1, LINK_COL).getValue())) return;
  markColumns_(sh, columnSpecs_()[TAB.schools], CONFIG.maxSchools, false);
}