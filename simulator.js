var simData = [];
var simParams = {};

function g(id) {
  return document.getElementById(id);
}

function runSim() {
  const startBalance = parseFloat(g("balance").value);
  const pct = parseFloat(g("percentage").value);
  const maxRisk = parseFloat(g("maxRisk").value);
  const maxTrades = parseInt(g("maxTrades").value);
  const target = parseFloat(g("target").value);
  const lossDays = parseInt(g("lossDays").value) || 0;
  const payout = 0.92;

  if ([startBalance, pct, maxRisk, maxTrades, target].some(isNaN)) {
    alert("Please fill all required fields.");
    return;
  }

  simParams = { startBalance, pct, maxRisk, maxTrades, target, lossDays };
  simData = [];

  let balance = startBalance;
  let peak = startBalance;
  let badDaysUsed = 0;
  const tbody = g("simBody");
  tbody.innerHTML = "";

  for (let day = 1; day <= 20; day++) {
    let dayStart = balance;
    let trades = 0;
    let dayResult = "Win";
    let stakeUsed = 0;
    let outcome = "";

    let expectedStake = Math.min(
      dayStart * (pct / 100),
      dayStart * (maxRisk / 100),
    );
    let expectedBalance = dayStart + expectedStake * payout;

    for (let t = 1; t <= maxTrades; t++) {
      let stake = balance * (pct / 100);
      let maxAllowed = balance * (maxRisk / 100);
      if (stake > maxAllowed) stake = maxAllowed;
      stakeUsed = stake;

      if (badDaysUsed < lossDays) {
        outcome += "L" + t + " ";
        dayResult = "Loss day";
        balance -= stake;
      } else {
        outcome += "W" + t + " ";
        balance += stake * payout;
        trades++;
        break;
      }
      trades++;
      if (balance <= 0) break;
    }

    if (badDaysUsed < lossDays) badDaysUsed++;

    const targetBalance = dayStart * (1 + target / 100);
    if (balance >= targetBalance) dayResult = "Target hit";
    if (balance > peak) peak = balance;
    const drawdown = ((peak - balance) / peak) * 100;

    const record = {
      day,
      startBalance: parseFloat(dayStart.toFixed(2)),
      stakeUsed: parseFloat(stakeUsed.toFixed(2)),
      outcome: outcome.trim(),
      expectedBalance: parseFloat(expectedBalance.toFixed(2)),
      actualBalance: parseFloat(balance.toFixed(2)),
      trades,
      result: dayResult,
      drawdownPct: parseFloat(drawdown.toFixed(2)),
    };
    simData.push(record);

    const badgeClass =
      dayResult === "Loss day"
        ? "badge-loss"
        : dayResult === "Target hit"
          ? "badge-target"
          : "badge-win";
    const row = document.createElement("tr");
    row.innerHTML =
      "<td>" +
      record.day +
      "</td>" +
      "<td>" +
      record.startBalance.toFixed(2) +
      "</td>" +
      "<td>" +
      record.stakeUsed.toFixed(2) +
      "</td>" +
      '<td style="font-size:12px;color:var(--color-text-secondary)">' +
      record.outcome +
      "</td>" +
      "<td>" +
      record.expectedBalance.toFixed(2) +
      "</td>" +
      '<td style="font-weight:500">' +
      record.actualBalance.toFixed(2) +
      "</td>" +
      "<td>" +
      record.trades +
      "</td>" +
      '<td><span class="badge ' +
      badgeClass +
      '">' +
      record.result +
      "</span></td>" +
      '<td style="color:' +
      (drawdown > 10
        ? "var(--color-text-danger)"
        : "var(--color-text-secondary)") +
      '">' +
      record.drawdownPct.toFixed(2) +
      "%</td>";
    tbody.appendChild(row);
    if (balance <= 0) break;
  }

  g("exportBtn").disabled = simData.length === 0;

  /* ── Summary stats ── */
  const last = simData[simData.length - 1];
  const pnl = last.actualBalance - simParams.startBalance;
  const ret = (pnl / simParams.startBalance) * 100;
  const maxDD = Math.max(
    ...simData.map(function (r) {
      return r.drawdownPct;
    }),
  );
  const wins = simData.filter(function (r) {
    return r.result !== "Loss day";
  }).length;
  const fmt = function (n) {
    return n.toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  };

  g("statFinal").textContent = fmt(last.actualBalance);
  g("statPnl").textContent = (pnl >= 0 ? "+" : "") + fmt(pnl);
  g("statPnl").className = "stat-value " + (pnl >= 0 ? "stat-pos" : "stat-neg");
  g("statReturn").textContent = (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%";
  g("statReturn").className =
    "stat-value " + (ret >= 0 ? "stat-pos" : "stat-neg");
  g("statDD").textContent = maxDD.toFixed(2) + "%";
  g("statDD").className =
    "stat-value " + (maxDD > 10 ? "stat-neg" : "stat-neutral");
  g("statWins").textContent = wins + " / " + simData.length;
  g("statWins").className = "stat-value stat-neutral";
  g("simStats").hidden = false;
}

function exportXLSX() {
  if (!simData.length) {
    alert("Run a simulation first.");
    return;
  }
  if (typeof XLSX === "undefined") {
    alert("Excel library not loaded yet, please try again.");
    return;
  }

  const wb = XLSX.utils.book_new();

  /* ── Sheet 1: Results ── */
  const headers = [
    "Day",
    "Start Balance",
    "Stake Used",
    "Outcome",
    "Expected Balance",
    "Actual Balance",
    "Trades",
    "Result",
    "Drawdown (%)",
  ];
  const rows = simData.map(function (r) {
    return [
      r.day,
      r.startBalance,
      r.stakeUsed,
      r.outcome,
      r.expectedBalance,
      r.actualBalance,
      r.trades,
      r.result,
      r.drawdownPct,
    ];
  });

  const wsData = [headers].concat(rows);
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  /* column widths */
  ws["!cols"] = [
    { wch: 6 },
    { wch: 16 },
    { wch: 13 },
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 8 },
    { wch: 12 },
    { wch: 14 },
  ];

  /* header row style */
  var range = XLSX.utils.decode_range(ws["!ref"]);
  for (var C = range.s.c; C <= range.e.c; C++) {
    var addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (!ws[addr]) continue;
    ws[addr].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "1A1A2E" } },
      alignment: { horizontal: "center" },
    };
  }

  /* data rows: number format + result colours */
  var resultColors = {
    Win: "C6EFCE",
    "Loss day": "FFC7CE",
    "Target hit": "CCFFEE",
  };
  var textColors = {
    Win: "276221",
    "Loss day": "9C0006",
    "Target hit": "0A5E43",
  };

  for (var R = 1; R <= rows.length; R++) {
    for (var C2 = 0; C2 <= 8; C2++) {
      var cellAddr = XLSX.utils.encode_cell({ r: R, c: C2 });
      if (!ws[cellAddr]) continue;
      ws[cellAddr].s = { alignment: { horizontal: "center" } };
      if (C2 === 1 || C2 === 2 || C2 === 4 || C2 === 5) {
        ws[cellAddr].z = "#,##0.00";
      }
      if (C2 === 8) {
        ws[cellAddr].z = '0.00"%"';
      }
      if (C2 === 7) {
        var result = rows[R - 1][7];
        ws[cellAddr].s = {
          alignment: { horizontal: "center" },
          font: { bold: true, color: { rgb: textColors[result] || "000000" } },
          fill: { fgColor: { rgb: resultColors[result] || "FFFFFF" } },
        };
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "Simulation Results");

  /* ── Sheet 2: Parameters ── */
  var paramsData = [
    ["Parameter", "Value"],
    ["Starting Balance", simParams.startBalance],
    ["Stake %", simParams.pct],
    ["Max Risk %", simParams.maxRisk],
    ["Max Trades/Day", simParams.maxTrades],
    ["Daily Target %", simParams.target],
    ["Simulated Bad Days", simParams.lossDays],
    ["Payout Rate", "92%"],
    ["Days Simulated", simData.length],
  ];
  var wsP = XLSX.utils.aoa_to_sheet(paramsData);
  wsP["!cols"] = [{ wch: 22 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsP, "Parameters");

  XLSX.writeFile(wb, "trading_simulation.xlsx");
}
