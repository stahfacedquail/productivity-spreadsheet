class Day {
  static list() {
    return [
      "Sunday", "Monday", "Tuesday", "Wednesday",
      "Thursday", "Friday", "Saturday",
    ];
  }

  /**
   * @params {number} dayIndex - 1 for the 1st; 2 for the 2nd; etc
   */
  constructor(dayIndex) {
    const days = Day.list();

    let length = 1;
    let goodAbbreviationFound = false;
    let possibleAbbreviations, possibleAbbreviation;
    while (!goodAbbreviationFound) {
      possibleAbbreviations = days.map(d => d.substring(0, length));
      possibleAbbreviation = days[dayIndex].substring(0, length);
      // If another day shares an abbreviation with this day, then the abbreviation
      // is not good enough.  Let's try a longer one.
      if (possibleAbbreviations.filter(abbr => abbr === possibleAbbreviation).length > 1) {
        length++;
      } else { // Otherwise, this is a good one; let's go!
        this.abbreviation = possibleAbbreviation;
        goodAbbreviationFound = true;
      }
    }
  }
}