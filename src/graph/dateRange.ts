import dayjs from "dayjs";

export function getDayRange(date: string) {
  const start = dayjs(date).startOf("day");
  if (!start.isValid()) throw new Error("Choose a valid day before loading your Daybook.");
  const end = start.add(1, "day");

  return {
    startZ: start.toDate().toISOString(),
    endZ: end.toDate().toISOString()
  };
}
