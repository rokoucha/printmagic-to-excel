import {
  readCSVObjects,
  writeCSV,
} from "https://deno.land/x/csv@v0.8.0/mod.ts";
import { z } from "https://deno.land/x/zod@v3.16.1/mod.ts";

const r = await Deno.open("./プリントマジック住所録データ.csv");

const PMAddress = z.object({
  分類: z.string(),
  フリガナ: z.string(),
  名前: z.string(),
  連名1: z.string(),
  敬称: z.string(),
  連名敬称1: z.string(),
  自宅郵便番号: z.string(),
  自宅住所: z.string(),
  会社郵便番号: z.string(),
  会社住所: z.string(),
  会社名: z.string(),
  部署: z.string(),
  役職: z.string(),
  "印刷住所（自宅/会社）": z.string(),
  "印刷する/しない": z.string(),
});
type PMAddress = z.infer<typeof PMAddress>;

const p: PMAddress[] = [];

for await (const obj of readCSVObjects(r)) {
  p.push(PMAddress.parse(obj));
}
r.close();

type EXAddress = {
  name: string;
  hurigana: string;
  sei: string;
  mei: string;
  renmei: string;
  keisho: string;
  shamei: string;
  busho: string;
  postalcode: string;
  address1: string;
  address2: string;
  tel: string;
  sent: string;
  received: string;
  remarks: string;
};
const e: EXAddress[] = p.map((p) => {
  const [[_, sei, mei]] = [...p.名前.matchAll(/([^\s]+)\s+([^\s]+)/g)];
  const [[__, address1, address2]] = [
    ...p.自宅住所.matchAll(/([^\s]+)(?:\s+([^\s]+))?/g),
  ];

  return {
    name: p.名前,
    hurigana: p.フリガナ,
    sei: sei ?? "",
    mei: mei ?? "",
    renmei: p.連名1,
    keisho: p.敬称,
    shamei: p.会社名,
    busho: p.部署,
    postalcode: p.自宅郵便番号,
    address1: address1 ?? "",
    address2: address2 ?? "",
    tel: "",
    sent: p["印刷する/しない"] === "する" ? "○" : "×",
    received: "0",
    remarks: "",
  };
});

const w = await Deno.open("./excel.csv", {
  write: true,
  create: true,
  truncate: true,
});

await writeCSV(w, [Object.keys(e[0]), ...e.map((er) => Object.values(er))]);

w.close();
