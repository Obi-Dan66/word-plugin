/* global module, window, process */

export function enableSoftHmr(): void {
  if (process.env.NODE_ENV === "production") {
    return;
  }

  if (module.hot) {
    module.hot.accept(() => {
      window.location.reload();
    });
  }
}
