import fs from "fs/promises";

export async function readFile() {
  try {
    const data = await fs.readFile("title.txt", "utf8");
    return data;
  } catch (error) {
    console.error(error);
  }
}
