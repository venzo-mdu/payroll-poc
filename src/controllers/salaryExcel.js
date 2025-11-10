export const parseSalarySheet = (sheetData) => {
  if (!sheetData || !sheetData.length) return {};

  const salaryData = {};
  const firstRow = sheetData[0];

  const allKeys = Object.keys(firstRow);
  const componentKey =
    allKeys.find((key) => key.toLowerCase().includes("blue ocea cost")) ||
    allKeys[0];

  const categoryKeys = allKeys.filter((key) => key !== componentKey);
  const categories = categoryKeys.map((key) => (firstRow[key] || "").trim());

  sheetData.slice(1).forEach((row) => {
    const field = row[componentKey];
    if (!field) return;

    categories.forEach((category, index) => {
      const colKey = categoryKeys[index];
      const value = row[colKey];

      if (!salaryData[category]) salaryData[category] = {};
      salaryData[category][field.trim()] = value ?? 0;
    });
  });

  return salaryData;
};
