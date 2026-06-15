export const formatMoney = (amount: number): string =>
  `${Math.round(amount).toLocaleString('ja-JP')}円`;

export const formatPercent = (value: number): string => `${value.toFixed(1)}%`;
