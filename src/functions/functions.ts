/* global clearInterval, CustomFunctions, Excel, setInterval */

/**
 * Increment a value every second using a streaming function
 * @customfunction
 * @param invocation Custom function handler
 */
export function incrementStreaming(invocation: CustomFunctions.StreamingInvocation<any[][]>): void {
  let result = 0;

  const timer = setInterval(() => {
    invocation.setResult([[result]]);

    result++;
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Increment a value every second using Excel.run
 * @customfunction
 * @requiresAddress
 * @param invocation Custom function handler
 */
export function increment(invocation: CustomFunctions.Invocation): any[][] {
  let result = 0;

  const [sheetId, cellId] = invocation.address.split("!");

  setInterval(() => {
    Excel.run((context) => {
      const cell = context.workbook.worksheets.getItem(sheetId).getRange(cellId);

      cell.values = [[result]];
      result++;

      return context.sync();
    });
  }, 1000);

  return [[result]];
}
