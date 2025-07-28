const SERVICE_ACCOUNT_EMAIL = 'score-reports@sat-score-reports.iam.gserviceaccount.com';
const dataLatestDate = '03/2025';

const cats = [
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

const satSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

const actSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};