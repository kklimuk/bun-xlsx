import { afterEach, expect, test } from "bun:test";
import {
	chmodSync,
	mkdirSync,
	mkdtempSync,
	readFileSync,
	realpathSync,
	rmSync,
	writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import pkg from "../package.json" with { type: "json" };

const enabled = process.env.XLSX_TEST_BINARY === "1" ? test : test.skip;
const temporaryDirectories: string[] = [];

afterEach(() => {
	for (const path of temporaryDirectories.splice(0))
		rmSync(path, { recursive: true, force: true });
});

async function run(
	binary: string,
	cwd: string,
	...args: string[]
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
	const child = Bun.spawn([binary, ...args], {
		cwd,
		stdout: "pipe",
		stderr: "pipe",
	});
	const [exitCode, stdout, stderr] = await Promise.all([
		child.exited,
		new Response(child.stdout).text(),
		new Response(child.stderr).text(),
	]);
	return { exitCode, stdout, stderr };
}

enabled("compiled binary works outside the repository", async () => {
	const directory = mkdtempSync(join(tmpdir(), "xlsx-binary-test-"));
	temporaryDirectories.push(directory);
	const copiedBinary = join(directory, "xlsx");
	await Bun.write(copiedBinary, Bun.file(resolve("dist/xlsx")));
	chmodSync(copiedBinary, 0o755);

	const help = await run(copiedBinary, directory, "--help");
	expect(help).toEqual(expect.objectContaining({ exitCode: 0, stderr: "" }));
	expect(help.stdout).toContain("xlsx <command> [options]");

	const version = await run(copiedBinary, directory, "--version");
	expect(version).toEqual({
		exitCode: 0,
		stdout: `xlsx ${pkg.version}\n`,
		stderr: "",
	});

	const upgrade = await run(copiedBinary, directory, "upgrade", "--dry-run");
	expect(upgrade.exitCode).toBe(0);
	expect(upgrade.stderr).toBe("");
	expect(JSON.parse(upgrade.stdout)).toEqual(
		expect.objectContaining({
			operation: "upgrade",
			dryRun: true,
			from: pkg.version,
			to: "latest",
			path: realpathSync(copiedBinary),
		}),
	);

	const fixture = resolve(
		"tests/fixtures/generated/22-formulas-without-cache.xlsx",
	);
	const read = await run(copiedBinary, directory, "read", fixture);
	expect(read.exitCode).toBe(0);
	expect(read.stderr).toBe("");
	expect(read.stdout).toContain("3,6,3,,87");
});

enabled(
	"installer validates before atomically replacing an existing binary",
	async () => {
		const directory = mkdtempSync(join(tmpdir(), "xlsx-installer-test-"));
		temporaryDirectories.push(directory);
		const release = join(directory, "release");
		const tools = join(directory, "tools");
		const prefix = join(directory, "bin");
		for (const path of [release, tools, prefix]) mkdirSync(path);
		const target = releaseTarget();
		const asset = join(release, target);
		const destination = join(
			prefix,
			process.platform === "win32" ? "xlsx.exe" : "xlsx",
		);
		const oldBinary = "#!/bin/sh\necho 'xlsx old'\n";
		writeFileSync(destination, oldBinary);
		chmodSync(destination, 0o755);
		const fakeCurl = join(tools, "curl");
		writeFileSync(
			fakeCurl,
			'#!/bin/sh\ncp "$FAKE_RELEASE/$(basename "$2")" "$4"\n',
		);
		chmodSync(fakeCurl, 0o755);

		await stageRelease(asset, release, "#!/bin/sh\nexit 9\n");
		const failed = await runInstaller(prefix, release, tools);
		expect(failed.exitCode).toBe(1);
		expect(failed.stderr).toContain("existing install is unchanged");
		expect(readFileSync(destination, "utf8")).toBe(oldBinary);

		await Bun.write(asset, Bun.file(resolve("dist/xlsx")));
		chmodSync(asset, 0o755);
		await writeChecksums(asset, release);
		const installed = await runInstaller(prefix, release, tools);
		expect(installed.exitCode).toBe(0);
		expect(installed.stdout).toContain(`xlsx ${pkg.version}`);
		const version = await run(destination, directory, "--version");
		expect(version).toEqual({
			exitCode: 0,
			stdout: `xlsx ${pkg.version}\n`,
			stderr: "",
		});
	},
);

function releaseTarget(): string {
	if (process.platform === "darwin")
		return process.arch === "arm64" ? "xlsx-darwin-arm64" : "xlsx-darwin-x64";
	if (process.platform === "linux")
		return process.arch === "arm64" ? "xlsx-linux-arm64" : "xlsx-linux-x64";
	return "xlsx-windows-x64.exe";
}

async function stageRelease(
	asset: string,
	release: string,
	contents: string,
): Promise<void> {
	await Bun.write(asset, contents);
	chmodSync(asset, 0o755);
	await writeChecksums(asset, release);
}

async function writeChecksums(asset: string, release: string): Promise<void> {
	const digest = new Bun.CryptoHasher("sha256")
		.update(await Bun.file(asset).bytes())
		.digest("hex");
	await Bun.write(
		join(release, "SHA256SUMS"),
		`${digest}  ${releaseTarget()}\n`,
	);
}

async function runInstaller(
	prefix: string,
	release: string,
	tools: string,
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
	const child = Bun.spawn(["sh", resolve("install.sh")], {
		cwd: resolve("."),
		stdout: "pipe",
		stderr: "pipe",
		env: {
			...process.env,
			FAKE_RELEASE: release,
			PATH: `${tools}:${process.env.PATH}`,
			PREFIX: prefix,
			VERSION: `v${pkg.version}`,
			REQUIRE_CHECKSUM: "1",
		},
	});
	const [exitCode, stdout, stderr] = await Promise.all([
		child.exited,
		new Response(child.stdout).text(),
		new Response(child.stderr).text(),
	]);
	return { exitCode, stdout, stderr };
}
