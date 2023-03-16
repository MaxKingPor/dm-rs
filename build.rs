use std::{env, fs, path::Path};
// use std::sync::OnceLock;

fn main() {
    let out_dir = env::var("OUT_DIR").unwrap();
    let path = Path::new(&out_dir).join("../../..");
    // 复制 文件到 输出目录
    fs::copy("dm.dll", path.join("dm.dll")).unwrap();

    // println!("cargo:rustc-link-arg=/DEF:DmReg.def");
    if let Ok(env) = env::var("CARGO_FEATURE_reg") {
        println!("CARGO_FEATURE_reg {}", env);

        std::process::Command::new(r#"lib.exe"#)
            .arg("/DEF:DmReg.def")
            .arg("/MACHINE:x86")
            .arg(format!(
                "/OUT:{}",
                Path::new(&out_dir).join("DmReg.lib").display()
            ))
            .spawn()
            .unwrap();

        println!("cargo:rerun-if-changed=DmReg.def");
        println!("cargo:rustc-link-search={}", out_dir);
        println!("cargo:rustc-link-lib=static=DmReg");
        fs::copy("DmReg.dll", path.join("DmReg.dll")).unwrap();
    }
}
