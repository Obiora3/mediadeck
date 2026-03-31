export const DEFAULT_SESSION_HOURS = 8;
export const DEFAULT_APP_SETTINGS = {
  vatRate: 7.5,
  roundToWholeNaira: false,
  sessionHours: DEFAULT_SESSION_HOURS,
  mpoTerms: [
    "Make-goods or low quality reproductions are unacceptable.",
    "Transmission/broadcast of wrong material is totally unacceptable and could be considered as non-compliance. Kindly contact the agency if material is bad or damaged.",
    "Payment will be made within 30 days after submission of invoice and COT, with compliance based on the client's tracking or monitoring report.",
    "Any change in rate must be communicated within 90 days from the expiration of this contract period.",
    "Please note that on no account should there be any change to this flighting without approval from the agency.",
    "Please acknowledge receipt and implementation of this order.",
    "As agreed, 1 complimentary spot will be run for every 4 paid spots to support the campaign objective."
  ],
};

export const mergeAppSettings = (saved = {}) => {
  const terms = Array.isArray(saved?.mpoTerms) && saved.mpoTerms.length ? saved.mpoTerms : DEFAULT_APP_SETTINGS.mpoTerms;
  return { ...DEFAULT_APP_SETTINGS, ...(saved || {}), mpoTerms: terms };
};
