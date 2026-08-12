function renderDashboard() {
  el.regionMeta.textContent = `${state.regions.length} rows`;
  el.districtMeta.textContent = `${state.districts.length} rows`;

  renderRankTable(el.regionRankHost, [
    [1, 1, 1],
    [2, 2, 3],
    [3, 3, 3],
    [4, 4, 4],
  ]);

  renderRankTable(el.districtRankHost, [
    [1, 1, 7],
    [2, 8, 14],
    [3, 15, 22],
    [4, 23, 29],
    [5, 30, 37],
  ]);

  renderWeightBand(el.regionWeightHost);
  renderWeightBand(el.districtWeightHost);

  buildGroupedTable(el.regionTable, state.regions, "Region");
  buildGroupedTable(el.districtTable, state.districts, "District");
  renderDistrictView();
}

function renderDistrictView() {
  const districtName = state.selectedDistrict;
  const districtRow = findDistrictRow(districtName);
  const storeRows = getDistrictStores(districtName);

  renderSummaryGrid(el.selectedDistrictSummary, districtRow);

  renderRankTable(el.storeRankHost, [
    [1, 1, 72],
    [2, 73, 144],
    [3, 145, 216],
    [4, 217, 288],
    [5, 289, 360],
  ]);

  renderWeightBand(el.storeWeightHost);

  buildGroupedTable(el.storeTable, storeRows, "District");
}
