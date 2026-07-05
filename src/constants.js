export const DEFAULT_CHANNELS = [
  // GROUP 1
  { id: 1, name: 'INV1 线包', group: 1 },
  { id: 2, name: 'INV1 磁芯', group: 1 },
  { id: 3, name: 'INV2 线包', group: 1 },
  { id: 4, name: 'INV2 磁芯', group: 1 },
  { id: 5, name: 'BST1 线包', group: 1 },
  { id: 6, name: 'BST1 磁芯', group: 1 },
  { id: 7, name: 'BST2 线包', group: 1 },
  { id: 8, name: 'BST2 磁芯', group: 1 },
  // GROUP 2
  { id: 9, name: 'TX1 线包', group: 2 },
  { id: 10, name: 'TX1 磁芯', group: 2 },
  { id: 11, name: 'TX2 线包', group: 2 },
  { id: 12, name: 'TX2 磁芯', group: 2 },
  { id: 13, name: 'TX3 线包', group: 2 },
  { id: 14, name: 'TX3 磁芯', group: 2 },
  { id: 15, name: 'EMI 继电器RY1', group: 2 },
  { id: 16, name: 'EMI 继电器RY3', group: 2 },
  // GROUP 3
  { id: 17, name: 'INV Q5', group: 3 },
  { id: 18, name: 'INV Q2', group: 3 },
  { id: 19, name: 'BST Q27', group: 3 },
  { id: 20, name: 'BST D86', group: 3 },
  { id: 21, name: 'H Q10', group: 3 },
  { id: 22, name: 'HQ17', group: 3 },
  { id: 23, name: 'L Q15', group: 3 },
  { id: 24, name: 'L Q8', group: 3 },
  // GROUP 4
  { id: 25, name: 'L Q16', group: 4 },
  { id: 26, name: 'L Q9', group: 4 },
  { id: 27, name: 'EMI L1线圈', group: 4 },
  { id: 28, name: 'EMI L1磁芯', group: 4 },
  { id: 29, name: 'EMI L3线圈', group: 4 },
  { id: 30, name: 'EMI L3磁芯', group: 4 },
  { id: 31, name: '腔体温度', group: 4 },
  { id: 32, name: '环境温度', group: 4 },
];

export const DEFAULT_CONFIG = {
  condition: '测试条件',
  pv: '功率',
  pvVoltage: '电压',
  pvCurrent: '电流',
  power: '--',
  voltageVal: '--',
  currentVal: '--',
  noteLabel: '备注',
  bottomNoteLabel: '底部总结备注',
  topNote: '--',
  styles: {
    condition: { bold: false, fontSize: 11 },
    pv: { bold: false, fontSize: 11 },
    pvVoltage: { bold: false, fontSize: 11 },
    pvCurrent: { bold: false, fontSize: 11 },
    power: { bold: true, fontSize: 11 },
    voltageVal: { bold: true, fontSize: 11 },
    currentVal: { bold: true, fontSize: 11 },
    noteLabel: { bold: false, fontSize: 11 },
    topNote: { bold: false, fontSize: 11 },
    bottomNoteLabel: { bold: false, fontSize: 11 },
    bottomNote: { bold: false, fontSize: 11 }
  }
};
