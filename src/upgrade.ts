import { basename, dirname } from "node:path";
import { parseArgs } from "node:util";
import installScript from "../install.sh" with { type: "text" };
import pkg from "../package.json" with { type: "json" };

const RELEASE_TAG = /^v?\d+\.\d+\.\d+(?:[-.][0-9A-Za-z.-]+)?$/;

export const UPGRADE_HELP = `xlsx upgrade — replace a standalone xlsx binary

Usage:
  xlsx upgrade [options]

Options:
  --to TAG     Install a specific release (for example v0.2.0); default: latest
  --dry-run    Report the intended replacement without changing anything
  -h, --help   Show this help

Only a standalone GitHub release binary upgrades itself. For an npm or Bun
global installation, use your package manager instead.
`;

export async function upgrade(args: string[]): Promise<number> {
	const parsed = parseArgs({
		args,
		allowPositionals: false,
		strict: true,
		options: {
			to: { type: "string" },
			"dry-run": { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
	});
	if (parsed.values.help) {
		process.stdout.write(UPGRADE_HELP);
		return 0;
	}
	const requested = parsed.values.to;
	if (requested && !RELEASE_TAG.test(requested))
		throw new Error("--to must be a release tag like v0.2.0");
	const target = requested
		? requested.startsWith("v")
			? requested
			: `v${requested}`
		: "latest";
	if (!isStandaloneBinary())
		throw new Error(
			`xlsx ${pkg.version} is package-manager-owned; upgrade it with: bun add -g @sageling/xlsx (or npm install -g @sageling/xlsx)`,
		);
	if (process.platform === "win32")
		throw new Error(
			"Windows cannot replace a running executable; install the new release manually",
		);
	const binaryPath = process.execPath;
	if (basename(binaryPath) !== "xlsx")
		throw new Error(`cannot self-upgrade a renamed binary: ${binaryPath}`);
	if (parsed.values["dry-run"]) {
		process.stdout.write(
			`${JSON.stringify({ operation: "upgrade", dryRun: true, from: pkg.version, to: target, path: binaryPath })}\n`,
		);
		return 0;
	}
	const child = Bun.spawn(["sh", "-s"], {
		stdin: new TextEncoder().encode(installScript),
		stdout: "inherit",
		stderr: "inherit",
		env: {
			...process.env,
			PREFIX: dirname(binaryPath),
			VERSION: target,
			REQUIRE_CHECKSUM: "1",
		},
	});
	const exitCode = await child.exited;
	if (exitCode !== 0)
		throw new Error(`upgrade failed; ${binaryPath} is unchanged`);
	const versionProcess = Bun.spawnSync([binaryPath, "--version"]);
	const installedVersion = versionProcess.stdout.toString().trim();
	if (versionProcess.exitCode !== 0 || !installedVersion.startsWith("xlsx "))
		throw new Error(
			`upgrade installed ${binaryPath}, but its version check failed`,
		);
	process.stdout.write(
		`Upgraded xlsx ${pkg.version} → ${installedVersion.slice(5)}\n`,
	);
	return 0;
}

function isStandaloneBinary(): boolean {
	return (
		import.meta.path.startsWith("/$bunfs/") || import.meta.path.includes("~BUN")
	);
}
