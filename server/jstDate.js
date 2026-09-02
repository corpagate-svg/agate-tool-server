const JST_TIME_ZONE = "Asia/Tokyo";

function jstDateParts(value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  if (!Number.isFinite(date.getTime())) return null;
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: JST_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const get = (type) => parts.find((part) => part.type === type)?.value;
  return { year: get("year"), month: get("month"), day: get("day") };
}

function jstDateString(value = new Date()) {
  const parts = jstDateParts(value);
  return parts ? `${parts.year}-${parts.month}-${parts.day}` : null;
}

function jstYearMonth(value = new Date()) {
  const parts = jstDateParts(value);
  return parts ? `${parts.year}-${parts.month}` : null;
}

function isSameJstDate(valueA, valueB) {
  const dateA = jstDateString(valueA);
  const dateB = jstDateString(valueB);
  return Boolean(dateA && dateB && dateA === dateB);
}

module.exports = { JST_TIME_ZONE, jstDateParts, jstDateString, jstYearMonth, isSameJstDate };
