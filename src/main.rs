






use dm_rs::Dmsoft;
use windows::Win32::System::Com;



#[allow(unused_labels)]

fn main() {
    unsafe{
        Com::CoInitializeEx(std::ptr::null(), Default::default()).unwrap();

        let dm = Dmsoft::new().unwrap();

        let s = dm.Ver().unwrap();
        println!("{}", s);
        let result = dm.SetPath("./").unwrap();

        println!("{:?}", result);

        
    }
    
    println!("#################################")
    
}

