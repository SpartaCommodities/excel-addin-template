/* global clearInterval, CustomFunctions, setInterval */

/**
 * Increment a value every second using a streaming function
 * @customfunction
 * @param invocation Custom function handler
 */
export function increment(invocation: CustomFunctions.StreamingInvocation<any[][]>): void {
  let result = 0;

  const timer = setInterval(() => {
    invocation.setResult([[result]]);

    result++;
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
