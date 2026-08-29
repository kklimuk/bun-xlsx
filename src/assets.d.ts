declare module "*.wasm" {
	const path: string;
	export default path;
}

declare module "*.sh" {
	const text: string;
	export default text;
}
