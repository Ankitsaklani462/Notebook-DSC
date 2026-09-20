import * as XLSX from 'xlsx'

export function downloadExcelReport({ rows, totals, categoryLosses, breakdownFields, numberValue, timeDifference, reportDate }) {
  const headers = ['Date', 'Line', 'Customer', 'Model', 'IE UPH', 'Shift', 'Available Time (Hr.)', 'PPC Plan', 'Actual Quantity', 'Gap/Excess', 'Ach %', 'Plan UPH', 'Actual UPH', 'L-Code', ...breakdownFields.map(([, label]) => label), 'Other', 'Total B/D Min.', 'Loss in Pcs', 'Reason', 'Root Cause', 'Action']
  
  // Helper function to extract loss minutes without double counting
  const getLossMinutesForCategory = (row, targetKey, targetLabel) => {
    let totalMinutes = 0
    
    // Check lossLogs array if available (prevents duplicate counting)
    if (row.lossLogs && Array.isArray(row.lossLogs) && row.lossLogs.length > 0) {
      row.lossLogs.forEach((loss) => {
        const cat = loss.reasonCategory || loss.category || ''
        if (cat === targetKey || cat === targetLabel) {
          totalMinutes += numberValue(loss.lossMinutes || timeDifference(loss.lossFrom, loss.lossTo))
        }
      })
    } else {
      // Fallback to single/main row entry category only if lossLogs is missing/empty
      const mainCat = row.reasonCategory || ''
      if (mainCat === targetKey || mainCat === targetLabel) {
        totalMinutes += numberValue(row.lossMinutes || timeDifference(row.lossFrom, row.lossTo))
      }
    }
    
    // Fallback to direct property if set and total is still 0
    if (totalMinutes === 0 && row[targetKey]) {
      totalMinutes += numberValue(row[targetKey])
    }
    
    return totalMinutes
  }

  const data = rows.map((row) => {
    // 1. Map breakdown dynamic fields
    const breakdownValues = breakdownFields.map(([key, label]) => {
      return getLossMinutesForCategory(row, key, label)
    })

    // 2. Map Machine (L-Code) loss minutes
    const lCodeMinutes = getLossMinutesForCategory(row, 'Machine (L-Code)', 'L-Code')
    
    // 3. Map Other loss minutes
    const otherMinutes = getLossMinutesForCategory(row, 'Other', 'Other') + numberValue(row.other)

    // Total Breakdown calculation
    const totalBreakdown = breakdownValues.reduce((sum, val) => sum + val, 0) + lCodeMinutes + otherMinutes
    
    const actual = numberValue(row.actual)
    const plan = numberValue(row.ppcPlan)
    const runMinutes = numberValue(row.runMinutes)
    const actualUph = runMinutes ? Math.round((actual * 60) / runMinutes) : 0
    const availableTimeHr = Number((runMinutes / 60).toFixed(2))
    const achPercentage = plan ? `${((actual / plan) * 100).toFixed(1)}%` : '0%'
    
    const losses = row.lossLogs?.length ? row.lossLogs : [{ reasonCategory: row.reasonCategory, lossFrom: row.lossFrom, lossTo: row.lossTo, lossMinutes: row.lossMinutes, reason: row.reason, rootCause: row.rootCause, action: row.action }]
    const reason = losses.filter((loss) => loss.reasonCategory || loss.reason).map((loss) => `${loss.reasonCategory ? `${loss.reasonCategory}: ` : ''}${loss.reason || ''}${loss.lossFrom && loss.lossTo ? ` (${loss.lossFrom}-${loss.lossTo},${loss.lossMinutes || timeDifference(loss.lossFrom, loss.lossTo)} min)` : ''}`).join('\n')
    const rootCause = losses.map((loss) => loss.rootCause).filter(Boolean).join('\n')
    const action = losses.map((loss) => loss.action).filter(Boolean).join('\n')
    
    const finalLCode = lCodeMinutes ? `${lCodeMinutes} min` : (row.lCode || '')

    return [
      row.date, row.line, row.customer, row.model, numberValue(row.ieUph), row.shift, 
      availableTimeHr, plan, actual, actual - plan, achPercentage, 
      Math.round(numberValue(row.planUph)), actualUph, finalLCode, 
      ...breakdownValues, 
      otherMinutes, totalBreakdown, 
      numberValue(row.ieUph) ? (totalBreakdown * 60) / numberValue(row.ieUph) : 0, 
      reason, rootCause, action
    ]
  })

  // Calculate Averages for Plan UPH and Actual UPH in Total Row
  const totalRowsCount = rows.length || 1
  const avgPlanUph = Math.round(rows.reduce((sum, row) => sum + Math.round(numberValue(row.planUph)), 0) / totalRowsCount)
  const avgActualUph = Math.round(rows.reduce((sum, row) => {
    const rm = numberValue(row.runMinutes)
    const act = numberValue(row.actual)
    return sum + (rm ? Math.round((act * 60) / rm) : 0)
  }, 0) / totalRowsCount)

  const totalAvailableTime = Number(rows.reduce((sum, row) => sum + (numberValue(row.runMinutes) / 60), 0).toFixed(2))
  const totalAchPercentage = totals.plan ? `${((totals.actual / totals.plan) * 100).toFixed(1)}%` : '0%'

  // Calculate totals for columns
  const totalBreakdownFieldsSum = breakdownFields.map(([key, label]) => {
    return rows.reduce((sum, row) => sum + getLossMinutesForCategory(row, key, label), 0)
  })

  const totalOtherSum = rows.reduce((sum, row) => sum + getLossMinutesForCategory(row, 'Other', 'Other') + numberValue(row.other), 0)
  const totalLCodeSum = rows.reduce((sum, row) => sum + getLossMinutesForCategory(row, 'Machine (L-Code)', 'L-Code'), 0)
  
  const grandTotalBreakdown = totalBreakdownFieldsSum.reduce((a, b) => a + b, 0) + totalLCodeSum + totalOtherSum

  const totalRow = [
    'TOTAL', '', '', '', '', '', 
    totalAvailableTime, totals.plan, totals.actual, totals.actual - totals.plan, totalAchPercentage, 
    avgPlanUph, avgActualUph, totalLCodeSum ? `${totalLCodeSum} min` : '', 
    ...totalBreakdownFieldsSum, 
    totalOtherSum, 
    grandTotalBreakdown, 
    '', '', '', ''
  ]

  const sheet = XLSX.utils.aoa_to_sheet([['PRODUCTION - NOTEBOOK', ...Array(headers.length - 1).fill('')], ['DAILY PRODUCTION REPORT', ...Array(headers.length - 1).fill('')], headers, ...data, totalRow])
  
  const headerRow = 2
  const dataStart = 3
  const totalRowIndex = dataStart + data.length
  
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 12 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } }]
  sheet['!cols'] = headers.map((header, index) => ({ wch: index >= 21 ? 28 : index >= 14 ? 11 : Math.max(10, Math.min(16, header.length + 2)) }))
  sheet['!rows'] = [{ hpt: 27 }, { hpt: 27 }, { hpt: 34 }, ...data.map(() => ({ hpt: 58 })), { hpt: 24 }]
  sheet['!freeze'] = { xSplit: 2, ySplit: 3 }
  
  const exportedRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true })
  const escapeHtml = (value) => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  
  const htmlRows = exportedRows.map((row, rowIndex) => {
    const isTitle = rowIndex < 2
    const isHeader = rowIndex === headerRow
    const isTotal = rowIndex === totalRowIndex
    const cells = row.map((value, column) => {
      const fill = isHeader ? (column >= 13 && column <= 24 ? '#ff0000' : '#5b9bd5') : isTotal ? '#d9ead3' : '#ffffff'
      const weight = isTitle || isHeader || isTotal ? 'bold' : 'normal'
      const align = column >= 21 ? 'left' : 'center'
      return `<td style="border:1px solid #808080;background:${fill};font-weight:${weight};text-align:${align};vertical-align:middle;white-space:pre-wrap;">${escapeHtml(value)}</td>`
    }).join('')
    return `<tr style="height:${isTitle ? 27 : isHeader ? 34 : 58}px">${cells}</tr>`
  }).join('')
  
  const html = `<html><head><meta charset="utf-8"><style>body{font-family:Arial;font-size:8pt}table{border-collapse:collapse}td{padding:4px;min-width:55px}tr:first-child td,tr:nth-child(2) td{font-size:15pt;height:27px}tr:nth-child(2) td{text-align:center}</style></head><body><table>${htmlRows}</table></body></html>`
  const downloadUrl = URL.createObjectURL(new Blob([html], { type: 'application/vnd.ms-excel' }))
  const downloadLink = document.createElement('a')
  downloadLink.href = downloadUrl
  downloadLink.download = `Daily-Report-${reportDate}.xls`
  downloadLink.click()
  URL.revokeObjectURL(downloadUrl)
}