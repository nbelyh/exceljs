// Backpressure bounds: pause once the queue is full, resume once it drains
// below the low mark. Pausing on a full queue is safe, because whatever the
// source still emits is already buffered, unlike pausing around every yield,
// which let slow consumers lose data (see below).
const HIGH_WATER_MARK = 64;
const LOW_WATER_MARK = 16;

module.exports = async function* iterateStream(stream) {
  const contents = [];
  stream.on('data', data => {
    contents.push(data);
    if (contents.length >= HIGH_WATER_MARK) {
      stream.pause();
    }
  });

  let resolveStreamEndedPromise;
  const streamEndedPromise = new Promise(resolve => (resolveStreamEndedPromise = resolve));

  let ended = false;
  stream.on('end', () => {
    ended = true;
    resolveStreamEndedPromise();
  });

  let error = false;
  stream.on('error', err => {
    error = err;
    resolveStreamEndedPromise();
  });

  while (!ended || contents.length > 0) {
    if (contents.length === 0) {
      stream.resume();
      // eslint-disable-next-line no-await-in-loop
      await Promise.race([once(stream, 'data'), streamEndedPromise]);
    } else {
      // Never pause around the yield itself. Sources such as unzipper's entry
      // stream keep working while paused and emit 'end' once their input is
      // exhausted; anything not yet delivered was then dropped silently, so
      // archive readers lost entries whenever the consumer awaited between
      // iterations. See the issue 1328 integration spec.
      const data = contents.shift();
      if (contents.length <= LOW_WATER_MARK) {
        stream.resume();
      }
      yield data;
    }
    if (error) throw error;
  }
  resolveStreamEndedPromise();
};

function once(eventEmitter, type) {
  // TODO: Use require('events').once when node v10 is dropped
  return new Promise(resolve => {
    let fired = false;
    const handler = () => {
      if (!fired) {
        fired = true;
        eventEmitter.removeListener(type, handler);
        resolve();
      }
    };
    eventEmitter.addListener(type, handler);
  });
}
