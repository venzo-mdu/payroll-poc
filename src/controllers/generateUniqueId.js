export function generateUniqueId() {
  const now = Date.now().toString(36);
  const rand = Math.random().toString(36).substring(2, 4);
  return (now + rand).substring(-6).toUpperCase().slice(-6);
}

export const requiredHeaders = [
  "S No",
  "Emp No",
  "Name",
  "Category",
  "DOJ",
  "Fixed Basic",
  "Fixed VDA",
  "HRA",
  "Gross",
  "Uniform",
  "Per day Cost",
  "Fixed Leave Wages",
  "Working days",
  "Man Days",
  "Allowance Days",
  "Gross",
  "Basic",
  "VDA",
  "HRA-1",
  "Leave Wages",
  "Bonus",
  "T.A. Allowance",
  "Total Earn",
  "Employee PF 12%",
  "Employee ESI 0.75%",
  "PT",
  "LWF",
  "Advances",
  "Canteen Deduction",
  "Other Deduction",
  "Uniform",
  "Total Deduc",
  "Net Pay",
  "Emplr PF 13%",
  "Emplr ESI 3.25%",
  "Empr LWF",
  "Uniform",
  "Total Cost",
  "Service Charge 5 %",
  "Cost",
  "New Bank Accounts",
  "IFSC code",
  "Salary Process",
];

export function findVal(jsonData, col) {
  return Object.values(jsonData[0]).findIndex((val) => val === col);
}
