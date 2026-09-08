const SERVICE_ACCOUNT_EMAIL = 'score-reports@sat-score-reports.iam.gserviceaccount.com';
const ADMIN_EMAIL = PropertiesService.getScriptProperties().getProperty('adminEmail');
const TEMPLATE_SHEETS_SS_ID = PropertiesService.getScriptProperties().getProperty('templateSheetsSsId');
const ACT_MASTER_DATA_SS_ID = PropertiesService.getScriptProperties().getProperty('actMasterDataSsId');
const RW_SCORE_LOOKUP = "=XLOOKUP($A$2&$A$3,'Rev sheet backend'!$T$71:$T$100,'Rev sheet backend'!U$71:U$100,)";
const MATH_SCORE_LOOKUP = "=XLOOKUP($A$2&$A$3,'Rev sheet backend'!$T$71:$T$100,'Rev sheet backend'!V$71:V$100,)";
const SKIPPED_IDS = ['080a7b51','1755eaf0','23974d6c','239d3535','271a5017','332e75bf','6fdddabf','714e4c10','a086c2cb','a266d876','a8fa749a','b168ce48','bcd924a5','c9708e7d','ca9dc00f','d222f7d9','d529c1ad','e060dd6b','fa07bcb6','00221c00','00460c13','0113152f','01989d77','01c8c433','03080769','03701ef3','0778b4ac','08b28c1a','08ff903e','0bcb4417','0c13dea9','0c622cfb','10cd0327','11a9f635','11df9b99','12d81fc1','14189fbb','145da981','16631d34','1724dac2','1782cdd7','190857f0','20a6a4ed','22105871','22b3da87','22e4d633','2584bcfb','27d9bb69','2903a041','296801d2','299c5303','2c50ed1a','2df7b582','2fdfe002','31ad8024','333b2b65','34d7bb25','3566120b','37a49687','3882ddf6','3f753a8e','4603d1f7','4c9a2aee','4eee64fa','50801257','54804e10','55688b3c','5ff1ba73','603755a5','61228830','62a18353','63c73b50','64e88c58','67667d72','69f031ab','6b49f5f1','6c9df5d1','6d44060a','6f5fc289','7254379e','787729be','7a1877be','7afdcca2','7fdba7ad','81da17d3','835d1ae6','84ece3f6','85439572','88bb0f6f','8de51658','9077be25','95dbdf51','9645f55e','96802cc0','97e2e364','97e5bf55','98fd50f2','9f1a0d91','a14eef71','a30567fd','a70cbc53','a9040290','aa5897b8','ad729337','adc8ea28','aecdb820','afec1a70','b0a525be','b1e8b87f','b4887dae','b74f676f','ba263620','c04e9136','c101fc44','c106b9f7','c14daa3c','c21df211','c4737d6a','c4d43991','c52652c9','c538954d','cac82f9b','d0fbf1ae','d2e0cba5','d3898d32','d3ca5d59','d46ac7e7','d5b9ed0d','d7f31e68','d8d1ecaa','db2e480a','db3ad406','de2c2f57','de3dd17d','df46a2ee','e1546fd6','e2829dd7','e459076b','e4f312c5','e677fa6c','e7247766','e7dc27dc','e929fe98','eef91a50','f07570bb','f38b40ac','f4b63a04','f4fd123c','f942646f','fbffb352','fce80a36','fdd9a360','ff3865b3','ff97fd53','16aa9758','36b6f8ba','3e28a622','5011b039','58ba7239','603c18ad','7268586c','744ee7d7','7b17f86a','9d2e7037','b78cd5df','c74df824','c78bb2b8','145337bc','1e562f24','290cdc2c','295a41f0','3122fc7b','441558e7','45df91ee','4b09f783','5639dd1a','5b6af6b1','6f6dfe3e','adb0c96c','c984f1a5','d9137a84','dcc4886a','fa2771d5','fc5ef8d3','fcb78856','3c03cbd8','4a090a46','7a8cb72a','e9f4521a']
dataLatestDate = '08/2025';
isUpdateAvailable = true;
areNewSatTestsAvailable = true;
isActSyncAvailable = false;


cats = [
  'Area and volume',
  'Boundaries',
  'Central ideas and details',
  'Circles',
  'Command of evidence',
  'Cross-text connections',
  'Distributions',
  'Equivalent expressions',
  'Form, structure, and sense',
  'Inferences',
  'Linear equations in one variable',
  'Linear equations in two variables',
  'Linear functions',
  'Linear inequalities',
  'Lines, angles, and triangles',
  'Models and scatterplots',
  'Nonlinear equations and systems',
  'Nonlinear functions',
  'Observational studies and experiments',
  'Percentages',
  'Probability',
  'Ratios, rates, proportions, and units',
  'Systems of linear equations',
  'Right triangles and trigonometry',
  'Sample statistics and margin of error',
  'Words in context',
  'Transitions',
  'Rhetorical synthesis',
  'Text, structure, and purpose',
];

const satSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

const actSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

const subjectData = [
  {
    'name': 'Reading & Writing',
    'rowOffset': 7,
  },
  {
    'name': 'Math',
    'rowOffset': 10
  }
]

