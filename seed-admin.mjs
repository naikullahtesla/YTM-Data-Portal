import { createClient } from "@supabase/supabase-js";
import * as fs from "node:fs";
import * as path from "node:path";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";

let SUPABASE_URL = process.env.VITE_SUPABASE_URL;
let SUPABASE_KEY = process.env.VITE_SUPABASE_ANON_KEY;

if (!SUPABASE_URL || !SUPABASE_KEY) {
  try {
    const envFile = fs.readFileSync(path.join(process.cwd(), ".env.local"), "utf8");
    for (const line of envFile.split("\n")) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;
      const eq = trimmed.indexOf("=");
      if (eq === -1) continue;
      const k = trimmed.slice(0, eq).trim();
      const v = trimmed.slice(eq + 1).trim();
      if (k === "VITE_SUPABASE_URL") SUPABASE_URL = v;
      if (k === "VITE_SUPABASE_ANON_KEY") SUPABASE_KEY = v;
    }
  } catch {}
}

if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error("Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY in .env.local");
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const rl = readline.createInterface({ input, output });
async function prompt(label) { return rl.question(label); }

async function main() {
  console.log("\n=== Create Admin User ===\n");

  const rawArgs = process.argv.slice(2);
  const force = rawArgs.includes("--force");
  const args = rawArgs.filter(a => a !== "--force");

  let email, name, password;

  if (args.length >= 2) {
    email = args[0].trim().toLowerCase();
    password = args[1];
    name = args[2] || "";
    console.log(`Email: ${email}`);
    console.log(`Name:  ${name || "(none)"}`);
  } else {
    email = (await prompt("Email: ")).trim().toLowerCase();
    name = (await prompt("Full name: ")).trim();
    password = await prompt("Password: ");
  }

  if (!email) { console.error("Email required."); process.exit(1); }
  if (!password) { console.error("Password required."); process.exit(1); }

  if (force) {
    await supabase.from("users").delete().eq("email", email);
  } else {
    const { data: existing } = await supabase.from("users").select("email").eq("email", email).single();
    if (existing) {
      console.error(`User ${email} already exists. Use --force to overwrite.`);
      process.exit(1);
    }
  }

  const { error } = await supabase.from("users").insert({
    email,
    password,
    full_name: name || null,
    role: "admin",
  });

  if (error) { console.error("Error:", error.message); process.exit(1); }

  console.log(`\nAdmin user created: ${email}`);
  rl.close();
}

main();
