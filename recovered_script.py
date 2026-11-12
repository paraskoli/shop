cat << 'EOF' > /tmp/generate_model.py
import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = '/home/aiteam/Desktop/shop'
PHOTO_DIR = os.path.join(BASE_DIR, 'Photos-1-001')
ORG_DIR = os.path.join(BASE_DIR, 'organized_shops')

# 1. Define categorization and files
CATEGORIES = {
    '01_Potato_Onion_Garlic': [
        ('IMG_20250706_172556.jpg', 'B-238_Narendra_Kumar_Naresh_Kumar_Potato_Onion_Garlic.jpg'),
        ('IMG_20250706_172602.jpg', 'B-235_Chaman_Khan_Co_Garlic_Onion_Chilli.jpg'),
        ('IMG_20250706_173711.jpg', 'B-256_Sarfaraz_Co_Potato_Onion_Veg.jpg'),
        ('IMG_20250706_174822.jpg', 'B-133_Haji_Dilshad_Sons_Potato_Onion_Crates.jpg'),
        ('IMG_20250706_175326.jpg', 'B-95_Rashid_Nawab_Saini_Trading_Garlic_Onion.jpg'),
        ('IMG_20250706_175328.jpg', 'B-96_Shyam_Sunder_Gulshan_Kumar_Potato_Onion_Wholesale.jpg'),
        ('IMG_20250706_175343.jpg', 'B-64_VIP_Fruit_Co_Dharamveer_Sons_Potato_Onion.jpg'),
        ('IMG_20250706_175420.jpg', 'B-67_Shri_Dauji_Trading_Co_Garlic_Potato_Onion.jpg'),
        ('IMG_20250706_175424.jpg', 'B-60_Sanchit_Trading_Co_Chipsona_Surya_Potatoes.jpg'),
        ('IMG_20250706_175428.jpg', 'B-70_Ravi_Kumar_Ramji_Das_Ginger_Garlic_Potato_Onion.jpg'),
        ('IMG_20250706_175430.jpg', 'B-71_Kasim_Tehseen_Hasanpur_Potato_Onion.jpg'),
        ('IMG_20250706_175444.jpg', 'B-56_Sambhal_Aloo_Company_Truck_Sacks.jpg'),
        ('IMG_20250706_175448.jpg', 'B-55_Vinod_Kumar_Kaushik_Potato_Onion_Wholesale.jpg'),
        ('IMG_20250706_175452.jpg', 'B-53_Jagdish_Saini_Suresh_Saini_Potato_Onion.jpg'),
        ('IMG_20250706_175524.jpg', 'B-17_Shyam_Sunder_Co_Nashik_MP_Onions.jpg'),
        ('IMG_20250706_175532.jpg', 'B-46_Hapur_Aloo_Sabzi_Bhandar_Panditji.jpg'),
        ('IMG_20250706_175605.jpg', 'B-16_Abid_Ali_Shafeeq_Garlic_Onion_Ginger.jpg'),
        ('IMG_20250706_175607.jpg', 'B-15_Chand_Sons_Onion_Garlic_Wholesale.jpg'),
        ('IMG_20250706_175609.jpg', 'B-14_Pandey_Trading_Co_Potato_Onion_Veg.jpg'),
        ('IMG_20250706_175618.jpg', 'B-11_Gopal_Das_Umesh_Kumar_Onion_Wholesale.jpg'),
        ('IMG_20250706_175620.jpg', 'B-9_Shahdara_Onion_Company_Kamal.jpg'),
        ('IMG_20250706_175635.jpg', 'B-7_B-6_Waseem_Co_Abdul_Aleem_Onion_Garlic.jpg'),
        ('IMG_20250706_175637.jpg', 'B-5_B-4_Kallu_Sons_Delhi_Veg_Co.jpg')
    ],
    '02_Tomato_Specialists': [
        ('IMG_20250706_174907.jpg', 'B-81_Haji_Habib_Sons_Special_In_Tomato_Wide.jpg'),
        ('IMG_20250706_174915.jpg', 'B-81_Haji_Habib_Sons_Special_In_Tomato_Sign.jpg'),
        ('IMG_20250706_175234.jpg', 'B-105_Ch_Pratap_Singh_Sons_Specialist_Tomato.jpg'),
        ('IMG_20250706_175239.jpg', 'B-88_Indian_Tomato_Company_Haji_Mukeem_Qureshi.jpg'),
        ('IMG_20250706_175256.jpg', 'A-71_B-4_Haji_Ikram_Tomato_Co_Crates_Yard.jpg'),
        ('IMG_20250706_175304.jpg', 'A-50_Neeraj_Naved_Aadhati_Tomato_Specialist.jpg'),
        ('IMG_20250706_175310.jpg', 'B-89_Arman_Trading_Co_Specialist_Tomato.jpg')
    ],
    '03_Green_Vegetables_Herbs': [
        ('IMG_20250706_173017.jpg', 'B-180_Rajdhani_Fruit_Veg_Cauliflower_Jackfruit.jpg'),
        ('IMG_20250706_173345.jpg', 'Lic_B-2129_Avinash_Kumar_Sons_Green_Veg.jpg'),
        ('IMG_20250706_173358.jpg', 'B-211_Naresh_Kumar_Sons_Dhaniya_Pudina_Specialist.jpg'),
        ('IMG_20250706_173604.jpg', 'B-242_Vedram_Baba_Bhindi_Parwal_Lauki_Specialist.jpg'),
        ('IMG_20250706_173652.jpg', 'B-249_Shri_Banke_Bihari_Fruit_Veg_Cucumber.jpg'),
        ('IMG_20250706_174813.jpg', 'B-126_Salmu_Sons_Cauliflower_Cabbage_Melons.jpg'),
        ('IMG_20250706_174816.jpg', 'B-125_B-126_Suryavanshi_Salmu_Green_Veg_Street.jpg'),
        ('IMG_20250706_174819.jpg', 'B-125_Suryavanshi_Rajput_Tomato_Capsicum_Brinjal.jpg'),
        ('IMG_20250706_174829.jpg', 'B-122_Prabhakar_Sons_Cabbage_Cauliflower.jpg'),
        ('IMG_20250706_174833.jpg', 'A-64_Meharban_Co_Radish_Carrot_Bhindi_Cucumber.jpg'),
        ('IMG_20250706_174851.jpg', 'B-115_Shoukee_Sons_Lucky_Aadhati_Seasonal_Veg.jpg')
    ],
    '04_Exotic_Frozen_Mushroom': [
        ('IMG_20250706_173631.jpg', 'A-171_JMD_Trading_Co_Mushroom_Exotic_Vegetables.jpg'),
        ('IMG_20250706_173842.jpg', 'B-275_Shyam_Lal_Surendra_Safal_Matar_Corn_Chaap.jpg')
    ],
    '05_Root_Vegetables': [
        ('IMG_20250706_173623.jpg', 'B-265_Sachin_Choudhary_Mooli_Arbi_Root_Veg.jpg')
    ],
    '06_Fruits_Bananas_Coconut': [
        ('IMG_20250706_173737.jpg', 'B-289_Suresh_Fruit_Veg_Fresh_Coconut_Pineapple.jpg'),
        ('IMG_20250706_173825.jpg', 'B-278_Bharat_Fruit_Co_Wide.jpg'),
        ('IMG_20250706_173827.jpg', 'B-278_Bharat_Fruit_Co_Signboard.jpg'),
        ('IMG_20250706_174635.jpg', 'B-294_B-295_Balaji_Fruit_Co_Bananas_Coconut.jpg'),
        ('IMG_20250706_174640.jpg', 'B-294_B-295_Balaji_Fruit_Co_Front_Carts.jpg'),
        ('IMG_20250706_175437.jpg', 'A-45_Om_Sai_Trading_Watermelon_Tarbooj_Potato.jpg')
    ],
    '07_General_Produce_Commission': [
        ('IMG_20250706_173640.jpg', 'B-264_Mahendra_Kumar_Mittal_Wholesale_Produce.jpg'),
        ('IMG_20250706_173641.jpg', 'B-264_Mahendra_Kumar_Mittal_Signboard_Close.jpg'),
        ('IMG_20250706_173655.jpg', 'B-251_Mahesh_Saini_Shivam_Saini_General_Veg.jpg'),
        ('IMG_20250706_173741.jpg', 'B-288_Brothers_Fruit_Vegetable_Wholesale.jpg'),
        ('IMG_20250706_173744.jpg', 'B-287_Shri_Banke_Bihari_Traders_Wholesale.jpg'),
        ('IMG_20250706_174826.jpg', 'B-135_Azad_Bhai_Co_Produce_Wholesale.jpg'),
        ('IMG_20250706_175317.jpg', 'B-101_RK_Traders_Ravi_Kohli_Wholesale.jpg')
    ],
    '08_Mandi_Walkthrough_Artifacts': [
        ('IMG_20250706_173812.jpg', 'Mandi_Walkthrough_Frame_173812_Camera_Blur.jpg')
    ]
}

# Create directories and symlinks
os.makedirs(ORG_DIR, exist_ok=True)
total_linked = 0

for cat, file_tuples in CATEGORIES.items():
    cat_dir = os.path.join(ORG_DIR, cat)
    os.makedirs(cat_dir, exist_ok=True)
    for src_file, dest_name in file_tuples:
        src_path = os.path.join(PHOTO_DIR, src_file)
        dest_path = os.path.join(cat_dir, dest_name)
        if os.path.exists(src_path):
            if os.path.lexists(dest_path):
                os.remove(dest_path)
            # Create relative symlink
            rel_src = os.path.relpath(src_path, cat_dir)
            os.symlink(rel_src, dest_path)
            total_linked += 1

print(f"Created organized directory structure with {total_linked} mapped photos.")
EOF
python3 /tmp/generate_model.py
