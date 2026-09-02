let tail = Promise.resolve();

function withSalesLock(fn) {
  const next = tail.then(fn, fn);
  tail = next.then(() => undefined, () => undefined);
  return next;
}

module.exports = { withSalesLock };
