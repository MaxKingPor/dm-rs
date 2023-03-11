use std::{env, path::{PathBuf, Path}, fs};
// use std::sync::OnceLock;

fn main(){
    let r = env::var("OUT_DIR").unwrap();
    let path = Path::new(&r).join("../../..");
    println!("OUT_DIR : {:?}", fs::canonicalize(&path));

    fs::copy("dm.dll", path.join("dm.dll")).unwrap();
    fs::copy("DmReg_b.dll", path.join("DmReg.dll")).unwrap();

}