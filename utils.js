const addHeadings = (people, headings) => {
    return people.map(personAsArray => {
      const personAsObj = {};
  
      headings.forEach((heading, i) => {
        personAsObj[heading] = personAsArray[i];
      });
  
      return personAsObj;
    });
  }
  
  const normalizeString = (value) => {
    return String(value || '')
      .trim()
      .toLowerCase();
  }
  
  const sheetValuesToObject = (sheetValues, headers) => {
    const headings = headers || sheetValues[0].map(normalizeString);
    let people = null;
    if (sheetValues) people = headers ? sheetValues : sheetValues.slice(1);
    const peopleWithHeadings = addHeadings(people, headings);
    return peopleWithHeadings;
  }
  
  module.exports = {
    addHeadings,
    normalizeString,
    sheetValuesToObject
  }
