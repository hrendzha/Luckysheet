import en from "./en";
import uk from "./uk";
import zh from "./zh";
import es from "./es";
import zh_tw from "./zh_tw";
import Store from "../store";

export const locales = { en, zh, es, zh_tw, uk };

function locale() {
  return locales[Store.lang];
}

export default locale;
