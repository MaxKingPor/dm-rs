






use dm_rs::Dmsoft;
use windows::Win32::System::Com;



#[allow(unused_labels)]

fn main() {
    unsafe{
        Com::CoInitializeEx(std::ptr::null(), Default::default()).unwrap();

        let dm = Dmsoft::new().unwrap();

        let s = dm.Ver().unwrap();
        println!("Ver: {}", s);
        let result = dm.SetPath("./").unwrap();

        println!("SetPath: {:?}", result);

        let result = dm.Ocr(0,0,2000,2000,"ffffff-000000",1.0);
        println!("Ocr: {:?}", result);
        
    }
    
    println!("#################################")
    
}

