import { expect, test } from "bun:test";
import { resolve } from "node:path";

test("installer rejects release-path injection before downloading", async () => {
	const child = Bun.spawn(["sh", resolve("install.sh")], {
		cwd: resolve("."),
		stdout: "pipe",
		stderr: "pipe",
		env: { ...process.env, VERSION: "../../other/releases/v9.9.9" },
	});
	const [exitCode, stdout, stderr] = await Promise.all([
		child.exited,
		new Response(child.stdout).text(),
		new Response(child.stderr).text(),
	]);
	expect(exitCode).toBe(1);
	expect(stdout).toBe("");
	expect(stderr).toContain("VERSION must be latest or a release tag");
});

test("installer never downloads and executes a script", async () => {
	const script = await Bun.file(resolve("install.sh")).text();
	expect(script).not.toMatch(/(?:curl|wget).*install\.sh.*\|.*sh/);
	expect(script).toContain('fetch_asset "$target" "$bin_tmp"');
	expect(script).toContain('"$bin_tmp" --version');
	expect(script.indexOf('"$bin_tmp" --version')).toBeLessThan(
		script.indexOf('mv "$bin_tmp"'),
	);
});
