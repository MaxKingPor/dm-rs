use std::{
    env, fs,
    path::{Path, PathBuf},
};
// use std::sync::OnceLock;

fn main() {
    let r = env::var("OUT_DIR").unwrap();
    let path = Path::new(&r).join("../../..");

    println!(
        "cargo:rustc-link-search={}",
        env::var("CARGO_MANIFEST_DIR").unwrap()
    );
    fs::copy("dm.dll", path.join("dm.dll")).unwrap();
    fs::copy("DmReg.dll", path.join("DmReg.dll")).unwrap();
}
