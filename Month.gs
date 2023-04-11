class Month {
  static list() {
    return [
      "January", "February", "March", "April",
      "May", "June", "July", "August",
      "September", "October", "November", "December",
    ];
  }

  /**
   * @param {number} monthIndex - zero-indexed (i.e. 0 for January)
   */
  constructor(monthIndex) {
    this.name = Month.list()[monthIndex];
    this.abbr = this.name.substring(0, 3);
  }
}
