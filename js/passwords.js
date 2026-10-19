// Password strength for encrypted saves. Anyone with a copy of a .wurd or protected .odt file can try passwords
// offline for as long as they like, so the password has to be hard to guess, not just long. zxcvbn (loaded on
// first use) estimates the guesses an attacker needs, counting the common passwords, words, names, dates and
// keyboard patterns they try first.
(function (Wurd) {
  'use strict';

  // About 2^64 guesses: out of reach even for a well-funded attacker, with each guess costing an Argon2id run
  const MIN_BITS = 64;
  // zxcvbn slows down on long input, and a password this long is decided by its start anyway
  const MAX_CHECKED = 100;
  // 32 symbols, so a random byte & 31 picks one without bias. No l, o, 0 or 1, which are easy to misread.
  const ALPHABET = 'abcdefghijkmnpqrstuvwxyz23456789';

  // { bits, warning }: bits is log2 of the estimated guesses; warning says what's weak, or is empty
  const strength = async (password, userInputs = []) => {
    await Wurd.loadScript('lib/zxcvbn.js');
    const result = window.zxcvbn(password.slice(0, MAX_CHECKED), userInputs.filter(Boolean));
    return { bits: Math.floor(result.guesses_log10 * Math.log2(10)), warning: result.feedback.warning || '' };
  };

  // Five groups of five random symbols (125 bits), e.g. "k7mqa-p9xre-..."
  const randomPassword = () => {
    const bytes = crypto.getRandomValues(new Uint8Array(25));
    const symbols = Array.from(bytes, (b) => ALPHABET[b & 31]);
    bytes.fill(0);
    return [0, 5, 10, 15, 20].map((i) => symbols.slice(i, i + 5).join('')).join('-');
  };

  // A random password that passes the check (zxcvbn could, very rarely, find a pattern in one)
  const suggest = async (userInputs = []) => {
    for (;;) {
      const password = randomPassword();
      if ((await strength(password, userInputs)).bits >= MIN_BITS) return password;
    }
  };

  Wurd.passwords = { MIN_BITS, strength, suggest };
})(window.Wurd = window.Wurd || {});
