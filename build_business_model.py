import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_PATH = '/home/aiteam/Desktop/shop/Vegetable_Omnichannel_Business_Model.xlsx'

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette
DARK_GREEN = "1E4620"
MED_GREEN = "2D6A4F"
LIGHT_GREEN = "D8F3DC"
PALE_GREEN = "F2F9F4"
ACCENT_ORANGE = "E76F51"
ACCENT_GOLD = "F4A261"
DARK_GRAY = "333333"
LIGHT_GRAY = "E5E5E5"
WHITE = "FFFFFF"

# Fonts
font_title = Font(name="Calibri", size=16, bold=True, color=WHITE)
font_subtitle = Font(name="Calibri", size=11, italic=True, color=LIGHT_GREEN)
font_section = Font(name="Calibri", size=13, bold=True, color=WHITE)
font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_subhdr = Font(name="Calibri", size=11, bold=True, color=DARK_GREEN)
font_bold = Font(name="Calibri", size=10, bold=True, color=DARK_GRAY)
font_regular = Font(name="Calibri", size=10, color=DARK_GRAY)
font_italic = Font(name="Calibri", size=9, italic=True, color="555555")
font_num = Font(name="Calibri", size=10, color=DARK_GRAY)
font_num_bold = Font(name="Calibri", size=10, bold=True, color=DARK_GREEN)

# Fills
fill_title = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
fill_section = PatternFill(start_color=MED_GREEN, end_color=MED_GREEN, fill_type="solid")
fill_header = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
fill_subhdr = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
fill_zebra = PatternFill(start_color=PALE_GREEN, end_color=PALE_GREEN, fill_type="solid")
fill_white = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
fill_accent = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

# Borders
thin_line = Side(border_style="thin", color="CCCCCC")
thick_top = Side(border_style="medium", color=DARK_GREEN)
double_bottom = Side(border_style="double", color=DARK_GREEN)
border_cell = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)
border_header = Border(left=thin_line, right=thin_line, top=thin_line, bottom=Side(border_style="medium", color="0D2818"))
border_total = Border(top=thick_top, bottom=double_bottom, left=thin_line, right=thin_line)

# Alignments
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

# -------------------------------------------------------------
# SHEET 1: VENDOR DIRECTORY
# -------------------------------------------------------------
ws1 = wb.create_sheet(title="Vendor_Directory")
ws1.views.sheetView[0].showGridLines = True

# Title banner
ws1.merge_cells("A1:N1")
ws1["A1"] = "GHAZIPUR APMC SUBZI MANDI — WHOLESALE VENDOR DIRECTORY & SOURCING CATALOG"
ws1["A1"].font = font_title
ws1["A1"].fill = fill_title
ws1["A1"].alignment = align_center

ws1.merge_cells("A2:N2")
ws1["A2"] = "Master catalog of verified mandi shops, aadhati contacts, produce specialties, procurement SOP ('How We Buy'), and retail/online selling models"
ws1["A2"].font = font_subtitle
ws1["A2"].fill = fill_title
ws1["A2"].alignment = align_center
ws1.row_dimensions[1].height = 28
ws1.row_dimensions[2].height = 20

headers1 = [
    "S.No", "Shop No", "Block", "Produce Category", "Trade / Firm Name", 
    "Proprietor / Aadhati Name", "Primary Phone", "Alternate Phone", 
    "APMC License No", "Specialist Produce / Varieties", "Standard Packaging Unit",
    "Organized Photo Reference", "How We Buy From Them (Procurement SOP)",
    "How We Sell In Store & Online (Retail Strategy)"
]

ws1.row_dimensions[3].height = 32
for col_num, h in enumerate(headers1, 1):
    c = ws1.cell(row=3, column=col_num)
    c.value = h
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

vendors_data = [
    (
        "B-238", "Block B", "Potato & Onion", "M/s Narendra Kumar Naresh Kumar",
        "Deepak Agarwal / Sudhir / Rakesh Singh", "8010204807", "9318370013 / 9718993110",
        "APMC Regd", "Potato (Sambhal/Agra), Onion (Nashik), Garlic", "50 kg Jute Sack / 40 kg Mesh Bag",
        "01_Potato_Onion_Garlic/B-238_Narendra_Kumar_Naresh_Kumar_Potato_Onion_Garlic.jpg",
        "Inspect sacks for rot, uniform sizing (45mm+), dry skin. Buy in 5-10 sack lots at 5:30 AM. Settle via weekly mandi account.",
        "Store: Ventilated sloped wooden bunkers, sell loose by kg. Online: 1kg & 2kg breathable kraft paper bags; combo with garlic."
    ),
    (
        "B-235", "Block B", "Garlic, Onion & Chilli", "M/s Chaman Khan & Company",
        "Chaman Aadhati / Sanni Aadhati / Ashraf", "9873459675", "8273786584 / 8273309768",
        "APMC Regd", "Garlic (Mandsaur/Ooty), Green Chilli, Onion", "20 kg Mesh Bag / 10 kg Crates",
        "01_Potato_Onion_Garlic/B-235_Chaman_Khan_Co_Garlic_Onion_Chilli.jpg",
        "Check garlic bulb tightness and moisture. Buy at 5:00 AM. Green chillies checked for crispness and stem retention.",
        "Store: Prominent basket near counter; complimentary chilli/dhaniya bundle. Online: 250g/500g punnets & zip-pouch packs."
    ),
    (
        "B-256", "Block B", "Potato & Onion", "M/s Sarfaraz & Co. / Raj & Co.",
        "Sarfaraz Bhai", "9818817730", "Mandi Desk",
        "APMC Regd", "Potatoes (New crop & cold store), Red Onions", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-256_Sarfaraz_Co_Potato_Onion_Veg.jpg",
        "Inspect for cuts from harvesting tools. Arrive 5:45 AM. Palledari ₹6/sack.",
        "Store: Bulk value discount (5kg bag offer). Online: Essential Pantry Combo (5kg Potato + 2kg Onion)."
    ),
    (
        "B-133", "Block B", "Potato & Onion", "M/s Haji Dilshad & Sons",
        "Mohd. Hamza", "8447212128", "Mandi Yard",
        "Lic No B-2282", "Potatoes, Onions, Green Veg crates", "Plastic Crates (25 kg) / Sacks (50 kg)",
        "01_Potato_Onion_Garlic/B-133_Haji_Dilshad_Sons_Potato_Onion_Crates.jpg",
        "Verified APMC license. Crates require ₹50 refundable security. Inspect crate contents from bottom to ensure honest grading.",
        "Store: Tiered shelf display, pre-graded into medium and large. Online: Premium grade washed potatoes in 1kg eco-pouches."
    ),
    (
        "B-95", "Block B", "Garlic & Onion", "M/s Rashid Nawab & Co. / Saini Trading",
        "Rashid Nawab / Saini Aadhati", "9718042500", "9810103595",
        "APMC Regd", "Garlic (Grade A large cloves), Red Onions", "10 kg / 25 kg Net Bags",
        "01_Potato_Onion_Garlic/B-95_Rashid_Nawab_Saini_Trading_Garlic_Onion.jpg",
        "Excellent source for premium large-clove garlic. Check for hollow cloves. Negotiate volume price on 5 net bags.",
        "Store: Hanging garlic braids and loose net display. Online: 250g/500g mesh bags for quick kitchen use."
    ),
    (
        "B-96", "Block B", "Potato & Onion", "M/s Shyam Sunder Gulshan Kumar",
        "Gulshan Kumar", "9811178648", "Mandi Desk",
        "APMC Regd", "Potato (3797 / Pukhraj), Red & Yellow Onions", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-96_Shyam_Sunder_Gulshan_Kumar_Potato_Onion_Wholesale.jpg",
        "Specialist order supplier. Order placement previous evening on WhatsApp; pickup at 5:00 AM ready loaded.",
        "Store: Daily staple counter. Online: Weekly vegetable subscription basket core item."
    ),
    (
        "B-64 / B-63", "Block B", "Potato & Onion", "M/s VIP Fruit Co. / Dharamveer & Sons / Aslam & Co.",
        "Dharamveer Saini / Aslam Bhai", "9811239088", "Mandi Yard",
        "APMC Regd", "Cold store Potatoes, Onions, Seasonal Fruit", "50 kg Sacks / Crates",
        "01_Potato_Onion_Garlic/B-64_VIP_Fruit_Co_Dharamveer_Sons_Potato_Onion.jpg",
        "Compare lot pricing with B-96. Check for soft rot in monsoon/humid weather. Settle by UPI.",
        "Store: Clean sorting into grade bins. Online: Listed under budget daily staples."
    ),
    (
        "B-67", "Block B", "Garlic, Potato & Onion", "M/s Shri Dauji Trading Co. / Abdul Salam",
        "Abdul Salam / Dauji Representative", "9818477610", "9871409212",
        "APMC Regd", "Garlic, Potato, Onion, Ginger", "25 kg / 50 kg Sacks",
        "01_Potato_Onion_Garlic/B-67_Shri_Dauji_Trading_Co_Garlic_Potato_Onion.jpg",
        "One-stop procurement for all aromatics & alliums. Arrive 5:15 AM to inspect arriving trucks from MP.",
        "Store: 'Aromatic Zone' with Ginger, Garlic & Onion grouped together. Online: Curated Tadka Box."
    ),
    (
        "B-60", "Block B", "Specialist Potatoes", "M/s Sanchit Trading Co. / Imran & Co.",
        "Sanchit / Imran Bhai", "9811428590", "9873114920",
        "APMC Regd", "Chipsona, Surya, Sugar-Free Potatoes", "50 kg Jute Sacks",
        "01_Potato_Onion_Garlic/B-60_Sanchit_Trading_Co_Chipsona_Surya_Potatoes.jpg",
        "High-end specialist varieties. Chipsona is low sugar, ideal for crispy cooking & diabetic diets. Inspect crispness.",
        "Store: Marketed as 'Sugar-Free Diabetic-Friendly Potatoes' at 40% premium. Online: Premium French-fry & baking potato SKU."
    ),
    (
        "B-70", "Block B", "Ginger, Garlic & Tubers", "M/s Ravi Kumar Ramji Das",
        "Ravi Kumar", "9811352467", "9871203495",
        "APMC Regd", "Ginger (Assam/Bangalore washed), Garlic, Onion", "20 kg Crates / 40 kg Sacks",
        "01_Potato_Onion_Garlic/B-70_Ravi_Kumar_Ramji_Das_Ginger_Garlic_Potato_Onion.jpg",
        "Inspect ginger for water soaking (overweight fraud) and fibrous quality. Premium unbleached ginger.",
        "Store: Dry wicker basket display; keep dry to prevent mold. Online: 200g sealed kraft punnets with freshness seal."
    ),
    (
        "B-71", "Block B", "Potato & Onion", "M/s Kasim Tehseen Hasanpur Wale",
        "Kasim / Tehseen Bhai", "8260390600", "7895563051",
        "APMC Regd", "Hasanpur/UP Potato, Red Onions", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-71_Kasim_Tehseen_Hasanpur_Potato_Onion.jpg",
        "Direct farm-truck arrivals from Hasanpur belt. High freshness, competitive mandi pricing.",
        "Store: Bulk value sales bins. Online: Monthly pantry staples pack."
    ),
    (
        "B-56", "Block B", "Direct Truck Potato", "M/s Sambhal Aloo Company",
        "Sambhal Mandi Representative", "9811289452", "9871158420",
        "APMC Regd", "Sambhal Fresh Potatoes (Direct Truckload)", "50 kg Jute Sacks",
        "01_Potato_Onion_Garlic/B-56_Sambhal_Aloo_Company_Truck_Sacks.jpg",
        "Direct procurement from truck tailgate. Lowest cost per kg, but require 10+ sack purchase. Arrive 4:30 AM.",
        "Store: High velocity daily seller; competitive with local hawkers. Online: 5kg family economy bag."
    ),
    (
        "B-55", "Block B", "Potato & Onion", "M/s Vinod Kumar Kaushik",
        "Vinod Kumar Kaushik", "9810711048", "9958431844",
        "APMC Regd", "Wholesale Potato & Onion, Garlic", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-55_Vinod_Kumar_Kaushik_Potato_Onion_Wholesale.jpg",
        "Reliable secondary source if Sambhal truck lots are sold out. Consistent medium grading.",
        "Store: Clean shelf bin stock. Online: Standard retail catalog SKU."
    ),
    (
        "B-53", "Block B", "Potato & Onion", "M/s Jagdish Saini & Suresh Saini",
        "Jagdish Saini / Suresh Saini", "9811342781", "9871448920",
        "APMC Regd", "Aloo-Pyaz Wholesale", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-53_Jagdish_Saini_Suresh_Saini_Potato_Onion.jpg",
        "Inspect skin condition and check for sprouting in warmer months. Buy fresh arrivals only.",
        "Store: Daily retail loose sales. Online: Essential daily vegetable pack."
    ),
    (
        "B-17", "Block B", "Specialist Onion", "M/s Shyam Sunder & Co. (Nashik / MP)",
        "Shyam Sunder", "9811124589", "9873105623",
        "APMC Regd", "Nashik Red Onion, Lasalgaon Special, MP Onion", "40 kg / 50 kg Mesh Sacks",
        "01_Potato_Onion_Garlic/B-17_Shyam_Sunder_Co_Nashik_MP_Onions.jpg",
        "Highest grade sweet red onions. Inspect for double-bulbs, moisture, and neck rot. Long shelf life (20+ days).",
        "Store: Hero onion display. Market as 'Original Nashik Red Onions'. Online: Premium 1kg & 3kg mesh bags."
    ),
    (
        "B-46", "Block B", "Specialist Potato", "M/s Hapur Aloo Sabzi Bhandar (Panditji)",
        "Panditji", "9811467290", "9871029340",
        "APMC Regd", "Hapur New Crop & Mountain Potatoes", "50 kg Sacks",
        "01_Potato_Onion_Garlic/B-46_Hapur_Aloo_Sabzi_Bhandar_Panditji.jpg",
        "Famous Hapur potato quality. High starch, quick boiling, delicious taste. Sourced by top restaurants.",
        "Store: Premium label 'Hapur Desi Aloo' with tasting notes. Online: Gourmet cooking potato SKU."
    ),
    (
        "B-16", "Block B", "Garlic, Onion & Ginger", "M/s Abid Ali Shafeeq / Mohd. Sharif Mohd. Salim",
        "Ashraf / Intezar Hussain", "9818896820", "9548644695",
        "APMC Regd", "Garlic, Onion, Ginger Wholesale", "25 kg Sacks / 20 kg Crates",
        "01_Potato_Onion_Garlic/B-16_Abid_Ali_Shafeeq_Garlic_Onion_Ginger.jpg",
        "Strong pricing on garlic and ginger combinations. Inspect ginger roots for freshness.",
        "Store: Daily kitchen staples corner. Online: Pre-weighed ginger-garlic combo pack."
    ),
    (
        "B-15", "Block B", "Onion & Garlic", "M/s Chand & Sons",
        "Chand Bhai / Sons", "9811382450", "9871294810",
        "APMC Regd", "Onion, Garlic, Seasonal Wholesale", "40 kg Mesh Bags",
        "01_Potato_Onion_Garlic/B-15_Chand_Sons_Onion_Garlic_Wholesale.jpg",
        "Inspect for uniform sizing (50mm size is ideal for household cooking). Fast loading.",
        "Store: Standard loose selling bins. Online: Household pantry standard."
    ),
    (
        "B-14", "Block B", "Potato & Seasonal Veg", "M/s Pandey Trading Co.",
        "Ishwar Chand Pandey", "9811478201", "9871502391",
        "APMC Regd", "Potato, Onion, Seasonal Vegetables", "50 kg Sacks / Crates",
        "01_Potato_Onion_Garlic/B-14_Pandey_Trading_Co_Potato_Onion_Veg.jpg",
        "Multi-commodity aadhati. Good for balancing leftover procurement shortages without walking entire mandi.",
        "Store: General shelf stock. Online: Standard grocery catalog."
    ),
    (
        "B-11", "Block B", "Specialist Onion", "M/s Gopal Das Umesh Kumar",
        "Umesh Kumar / Gopal Das", "9873309304", "9810339767",
        "APMC Regd", "Wholesale Onion Specialists (Grade A)", "50 kg Mesh Bags",
        "01_Potato_Onion_Garlic/B-11_Gopal_Das_Umesh_Kumar_Onion_Wholesale.jpg",
        "Premier onion specialist. Always offers well-cured onions with zero internal black mold. Check 3 sample bags.",
        "Store: Premium shelf display. Online: Grade A certified 1kg zip-lock bags."
    ),
    (
        "B-9", "Block B", "Specialist Onion", "M/s Shahdara Onion Company",
        "Kamal", "9811381165", "Mandi Yard",
        "APMC Regd", "Wholesale Onion Commission Agents", "50 kg Mesh Bags",
        "01_Potato_Onion_Garlic/B-9_Shahdara_Onion_Company_Kamal.jpg",
        "High volume trader with deep supplier lines. Ideal for bulk purchases during price escalation cycles.",
        "Store: Bulk value discount offers. Online: Bulk 5kg economy bag."
    ),
    (
        "B-7 / B-6", "Block B", "Onion & Garlic", "M/s Waseem & Co. / Abdul Aleem & Co.",
        "Waseem Bhai / Abdul Aleem", "9897007632", "9711726465",
        "APMC Regd", "Wholesale Onions, Garlic & Tubers", "40 kg / 50 kg Sacks",
        "01_Potato_Onion_Garlic/B-7_B-6_Waseem_Co_Abdul_Aleem_Onion_Garlic.jpg",
        "Competitive quotes on onion crates. Inspect bottom layers of sacks. Settle via cash/UPI on delivery.",
        "Store: Everyday staple bins. Online: Standard 1kg offering."
    ),
    (
        "B-5 / B-4", "Block B", "Produce & Vegetables", "M/s Kallu & Sons / Delhi Vegetable Co.",
        "Kallu Bhai / Delhi Veg Manager", "9811294801", "9871034821",
        "APMC Regd", "Wholesale Vegetables, Potato, Onion", "Crates & Sacks",
        "01_Potato_Onion_Garlic/B-5_B-4_Kallu_Sons_Delhi_Veg_Co.jpg",
        "Dual shop operation with high crate turnover. Check crate freshness at 5:00 AM.",
        "Store: General vegetable display. Online: Multi-vegetable bundles."
    ),
    (
        "B-81", "Block B", "Tomato Specialist", "M/s Haji Habib & Sons (Special in Tomato)",
        "Mohd. Habib / Mohd. Arshad", "9312211902", "9818239328 / 9313452252",
        "APMC Regd", "Hybrid & Desi Tomatoes (Shivaji, Himsona)", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/B-81_Haji_Habib_Sons_Special_In_Tomato_Sign.jpg",
        "Premier tomato specialist in Ghazipur. Buy 2 crates: 1 fully ripe red (for same-day sales) and 1 semi-ripe pink (for next-day). Check firmness.",
        "Store: Cushion crates with bubble wrap; display on sloped tiers. Online: 500g & 1kg puncture-free ventilated punnets."
    ),
    (
        "B-105", "Block B", "Tomato Specialist", "M/s Ch. Pratap Singh & Sons",
        "Ravi Choudhary", "9718755113", "9810842738",
        "APMC Regd", "Specialist in Tomato (Bangalore / Nasik Tomato)", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/B-105_Ch_Pratap_Singh_Sons_Specialist_Tomato.jpg",
        "Excellent firm hybrid tomatoes with thick walls (longer shelf life: 5-7 days). Arrive 4:15 AM.",
        "Store: Firm slicing tomatoes for salads and sandwiches. Online: Premium salad tomato packs."
    ),
    (
        "B-88", "Block B", "Tomato Specialist", "M/s Indian Tomato Company",
        "Haji Mukeem Qureshi", "9811276020", "9871542100",
        "APMC Regd", "Tomatoes (All varieties: Desi, Cherry, Hybrid)", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/B-88_Indian_Tomato_Company_Haji_Mukeem_Qureshi.jpg",
        "High volume specialist. Inspect crate corners for squashed fruit. Negotiate rate on 3+ crates.",
        "Store: Separate Desi (tangy, curry cut) from Hybrid (firm, salad cut). Online: Clearly labelled Desi vs Hybrid SKU."
    ),
    (
        "A-71 / B-4", "Block A & B", "Tomato Specialist", "M/s Haji Ikram Tomato Co. & ABC Tomato Co.",
        "Rinky / Nitin", "9871710920", "9810052341",
        "APMC Regd", "Wholesale Tomatoes (Yard Crates)", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/A-71_B-4_Haji_Ikram_Tomato_Co_Crates_Yard.jpg",
        "Large open yard operations. Direct truck unload. Crate deposit ₹100 required. Fast turn-around.",
        "Store: Bulk weekend discount deals. Online: Puree & Gravy Tomato Combo (2kg bulk pack)."
    ),
    (
        "A-50", "Block A", "Tomato & Capsicum", "M/s Neeraj Bhai & Naved Bhai Aadhati",
        "Neeraj Bhai / Naved Bhai / Anil Rajput", "9868503495", "9958543606",
        "APMC Regd", "Specialist Tomato, Shimla Mirch (Green/Colored)", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/A-50_Neeraj_Naved_Aadhati_Tomato_Specialist.jpg",
        "Combined sourcing of tomatoes and capsicum. Check capsicum for firm walls and bright green gloss. 4:30 AM arrival.",
        "Store: Vibrant color block display (Red tomato next to Green capsicum). Online: 'Stir-Fry Kit' bundle."
    ),
    (
        "B-89", "Block B", "Tomato & Veg Specialist", "M/s Arman Trading Co. / Arvind Kumar Trivedi",
        "Arvind Kumar Trivedi / Arman Bhai", "9313452252", "9811342618",
        "APMC Regd", "Tomatoes, Seasonal Veg Wholesale", "Plastic Crates (25 kg)",
        "02_Tomato_Specialists/B-89_Arman_Trading_Co_Specialist_Tomato.jpg",
        "Secondary backup for high-demand tomato days. Test ripeness by gentle thumb pressure.",
        "Store: Shelf restock by 1:00 PM for evening rush. Online: Standard grocery catalog."
    ),
    (
        "B-211", "Block B", "Herbs & Leafy Greens", "M/s Naresh Kumar & Sons (Dhaniya Wale)",
        "Naresh Kumar / Deepu Kumar / Ankit Kumar", "9811467479", "9871714494",
        "Lic No Regd", "Fresh Dhaniya (Coriander), Pudina (Mint), Methi", "Jute Bundles / Bamboo Baskets (20 kg)",
        "03_Green_Vegetables_Herbs/B-211_Naresh_Kumar_Sons_Dhaniya_Pudina_Specialist.jpg",
        "Critical vendor! Buy at 4:00 AM sharp. Inspect root moisture and yellow leaf ratio. Shake out excess mud. Transport in moist cloth.",
        "Store: Keep on hydration misting rack; trim stem roots; mist every 45 mins. Online: 100g/250g moisture-lock sealed paper bags."
    ),
    (
        "B-242", "Block B", "Green Vegetables", "M/s Vedram Baba Aadhati",
        "Sanjay Kashyap / Babu Ram Kashyap", "9718806021", "8750899454",
        "APMC Regd", "Bhindi (Okra), Parwal (Pointed Gourd), Lauki", "Plastic Crates (20 kg) / Gunny bags",
        "03_Green_Vegetables_Herbs/B-242_Vedram_Baba_Bhindi_Parwal_Lauki_Specialist.jpg",
        "Perform tip-snap test on okra (must snap crisply without bending). Parwal should be firm and seed-light. Lauki should be tender and scratch-pale.",
        "Store: Single layer display on green astroturf; avoid stacking heavy veg on okra. Online: 500g ventilated brown kraft pouches."
    ),
    (
        "B-180", "Block B", "Cruciferous & Jackfruit", "M/s Rajdhani Fruit & Vegetable",
        "Ram Naresh Mishra Aadhati", "9011926007", "9654663341",
        "APMC Regd", "Cauliflower (Gobhi), Cabbage, Kathal (Jackfruit)", "Mesh Sacks (30 kg) / Open Crates",
        "03_Green_Vegetables_Herbs/B-180_Rajdhani_Fruit_Veg_Cauliflower_Jackfruit.jpg",
        "Select cauliflower with tight white curds and fresh green jacket leaves (protects curd). Jackfruit should have fresh milky sap.",
        "Store: Trim leaves to 1 inch; display heads up. Online: Whole head (400-600g) or pre-cut florets; vacuum-packed peeled kathal."
    ),
    (
        "Lic_B-2129", "Block B", "Green Vegetables", "M/s Avinash Kumar & Sons",
        "Avinash Kumar", "Mandi Yard", "APMC Desk",
        "Lic No B-2129", "Seasonal Green Vegetables & Pods", "Crates (20 kg)",
        "03_Green_Vegetables_Herbs/Lic_B-2129_Avinash_Kumar_Sons_Green_Veg.jpg",
        "Check beans (French beans/Sem) for stringiness and seed bulges. Buy crisp young pods.",
        "Store: Sloped wicker trays. Online: 'Fresh Greens' recipe kit."
    ),
    (
        "B-122", "Block B", "Brassica Specialists", "M/s Prabhakar & Sons",
        "D.S. Prabhakar / Dalchand", "9212159894", "9718586335 / 9873747611",
        "APMC Regd", "Cabbage (Patta Gobhi), Cauliflower (Phool Gobhi)", "Loose Crates / Sacks",
        "03_Green_Vegetables_Herbs/B-122_Prabhakar_Sons_Cabbage_Cauliflower.jpg",
        "Inspect cabbage weight-to-size ratio (dense heads are higher quality). Remove loose outer 2 leaves.",
        "Store: Cool shelf placement. Online: Unit price per head with weight range guarantee."
    ),
    (
        "B-125", "Block B", "Fruiting Vegetables", "M/s Suryavanshi Rajput Trading Co. (SRTC)",
        "Satish Chauhan", "9871147040", "9212003889",
        "APMC Regd", "Brinjal (Baingan - Bharta & Small), Capsicum, Chillies", "Plastic Crates (20 kg)",
        "03_Green_Vegetables_Herbs/B-125_Suryavanshi_Rajput_Tomato_Capsicum_Brinjal.jpg",
        "Inspect brinjal skin gloss (dull skin means overripe/seedy) and stem thorns. Check for borer insect holes.",
        "Store: Soft fabric pad beneath brinjals to prevent bruising. Online: Bharta Special kit (Large baingan + tomato + garlic + green chilli)."
    ),
    (
        "B-126", "Block B", "Vegetables & Melons", "M/s Salmu & Sons / Pappu Mewati",
        "Pappu Mewati / Nazim Aadhati", "9810573921", "9718423610",
        "APMC Regd", "Cauliflower, Cabbage, Melons (Kharbooza)", "Crates & Net Bags",
        "03_Green_Vegetables_Herbs/B-126_Salmu_Sons_Cauliflower_Cabbage_Melons.jpg",
        "Excellent seasonal melons during summer and brassicas during winter. Tap melons for hollow resonance.",
        "Store: Fruit & vegetable boundary display. Online: Summer fruit & salad refreshers."
    ),
    (
        "A-64", "Block A", "Root & Salad Veg", "M/s Meharban & Company",
        "Meharban Bhai", "9811497250", "9873521094",
        "APMC Regd", "Mooli (Radish), Gajar (Carrot), Bhindi, Kheera", "Gunny Sacks / Crates (25 kg)",
        "03_Green_Vegetables_Herbs/A-64_Meharban_Co_Radish_Carrot_Bhindi_Cucumber.jpg",
        "Check mooli and gajar crunch by snapping one sample root. Green leafy tops should be vibrant (indicates freshly harvested 12 hrs ago).",
        "Store: Wash roots thoroughly; display with trimmed green tops. Online: Daily Salad Box (Radish, Carrot, Cucumber, Lemon, Green Chilli)."
    ),
    (
        "B-115", "Block B", "Seasonal Produce", "M/s Shoukee & Sons (Lucky Aadhati)",
        "Lucky Aadhati / Shoukee", "7065586791", "9990525658 / 8750899454",
        "APMC Regd", "Seasonal Gourds, Sponge Gourd (Tori), Bitter Gourd (Karela)", "Bamboo Baskets (15 kg)",
        "03_Green_Vegetables_Herbs/B-115_Shoukee_Sons_Lucky_Aadhati_Seasonal_Veg.jpg",
        "Inspect karela for dark green warts (sharper taste, premium) and tori for tender skin. Avoid yellowing gourd ends.",
        "Store: Arranged in shallow wood crates. Online: 500g portions in eco-friendly ventilated paper bags."
    ),
    (
        "B-249", "Block B", "Cucumber & Gourds", "M/s Shri Banke Bihari Fruit & Veg",
        "Subhash Bhai / Shashi Bhai", "8076527926", "9810879260",
        "APMC Regd", "Kheera (Desi & English Cucumber), Kakdi, Gourds", "Plastic Crates (25 kg)",
        "03_Green_Vegetables_Herbs/B-249_Shri_Banke_Bihari_Fruit_Veg_Cucumber.jpg",
        "Ensure cucumbers are straight, firm, and seedless. English seedless cucumbers must have smooth wax-free skin.",
        "Store: Upright in hydration misting section. Online: English cucumber individual shrink-wrap; Desi cucumber 500g pack."
    ),
    (
        "A-171", "Block A", "Exotics & Mushrooms", "M/s JMD Trading Co.",
        "JMD Sourcing Manager", "9811224858", "9871766624",
        "APMC Regd", "Button Mushroom, Oyster Mushroom, Baby Corn, Broccoli, Safal Peas", "10-Punnet Corrugated Boxes (200g x 10)",
        "04_Exotic_Frozen_Mushroom/A-171_JMD_Trading_Co_Mushroom_Exotic_Vegetables.jpg",
        "High-margin exotics! Button mushrooms must have closed caps (no black gills showing) and dry snow-white color. Keep in insulated vehicle box during transit.",
        "Store: Displayed inside commercial glass-door chiller at 4°C - 8°C. Online: 200g punnets; Exotics Stir-Fry Basket (Broccoli + Baby Corn + Mushroom)."
    ),
    (
        "B-275", "Block B", "Frozen Foods & Dairy", "M/s Shyam Lal Surendra Kumar (Happy Bajaj)",
        "Happy Bajaj / Surendra Kumar", "9643224014", "8744817289",
        "APMC Regd", "Safal Frozen Green Peas, Sweet Corn, Fresh Soya Chaap, Paneer", "1 kg / 5 kg Frozen Food Service Packs",
        "04_Exotic_Frozen_Mushroom/B-275_Shyam_Lal_Surendra_Safal_Matar_Corn_Chaap.jpg",
        "Check freezer temperatures (-18°C for peas/corn). Ensure no ice clump recrystallization (indicates defrosting cycle). Fresh soya chaap delivered daily.",
        "Store: Deep freezer section beside checkout counter. Online: Frozen pantry category; bundled with paneer & butter for 'Matar Paneer Meal Kit'."
    ),
    (
        "B-265", "Block B", "Root Specialists", "M/s Sachin Choudhary",
        "Sachin Choudhary", "7983022735", "Mandi Yard",
        "Lic No B-65", "Mooli (Radish), Gajar (Carrot), Arbi (Colocasia), Shalgam", "Jute Sacks (40 kg)",
        "05_Root_Vegetables/B-265_Sachin_Choudhary_Mooli_Arbi_Root_Veg.jpg",
        "Arbi must be hard with dry fibrous skin (no soft squishy roots). Mooli washed and sorted. Buy at 5:00 AM.",
        "Store: Sloped rustic wooden bins; keep arbi completely dry. Online: Peeled & sanitized Arbi (300g ready-to-cook pack at 50% margin)."
    ),
    (
        "B-294 / B-295", "Block B", "Bananas & Coconut", "M/s Balaji Fruit Company",
        "Jay Prakash / Suresh Yadav / Gavendra Singh / Nitin Dhingra", "9910276002", "8287365066 / 8368171037",
        "APMC Regd", "Bananas (Robusta / G9 - AC Ripened), Fresh Green Coconuts (Daak)", "Banana Crates (15 kg) / Coconut Bags (25 pcs)",
        "06_Fruits_Bananas_Coconut/B-294_B-295_Balaji_Fruit_Co_Bananas_Coconut.jpg",
        "Bananas ripened in ethylene AC chambers (Grade 4 green-yellow tip for 3-day shelf life). Coconuts shaken for water content.",
        "Store: Suspended banana hanging hooks (prevents bruising); coconut water cutting station at store front. Online: 'Dozen Bananas' & 'Tender Coconut'."
    ),
    (
        "B-289", "Block B", "Fresh Coconut & Fruit", "M/s Suresh Fruit & Vegetable",
        "Suresh Bhai", "9818314144", "9213568822",
        "APMC Regd", "Fresh Nariyal (Dry & Water Coconuts), Pineapple, Citrus", "Gunny Sacks / Wooden Crates",
        "06_Fruits_Bananas_Coconut/B-289_Suresh_Fruit_Veg_Fresh_Coconut_Pineapple.jpg",
        "Large dry puja coconuts and sweet water tender coconuts. Event & festival wholesale supplier.",
        "Store: Front porch temple / festival corner display. Online: Puja essentials bundle & immunity fruit pack."
    ),
    (
        "B-278", "Block B", "Seasonal Fruits", "M/s Bharat Fruit Company",
        "Bharat Bhai", "9810238814", "Mandi Desk",
        "APMC Regd", "Apples (Shimla/Kashmir), Pears, Citrus, Seasonal Fruits", "Telescopic Corrugated Cartons (18-20 kg)",
        "06_Fruits_Bananas_Coconut/B-278_Bharat_Fruit_Co_Signboard.jpg",
        "Inspect tray pack layers for transit bruising. Check pressure firmness. Buy 2-3 cartons 3 times weekly.",
        "Store: Eye-level central island display with foam net protectors. Online: Grade A fruit packs (4-piece & 1kg packs)."
    ),
    (
        "A-45", "Block A", "Watermelons & Hasanpur", "M/s Om Sai Trading Company",
        "Nawab Bhai / Zahid Bhai (Hasanpur Wale)", "9811571249", "Mandi Yard",
        "APMC Regd", "Watermelons (Tarbooj - Black & Striped), Hasanpur Potatoes", "Loose Truckload / Net Crates",
        "06_Fruits_Bananas_Coconut/A-45_Om_Sai_Trading_Watermelon_Tarbooj_Potato.jpg",
        "Check watermelons for yellow ground spot (sun ripened) and deep dull thump sound. Avoid shiny green rind.",
        "Store: Cut half-watermelon wrapped in cling film on ice; whole melons at entrance. Online: Whole melon (2.5-3.5kg) with sweetness guarantee."
    ),
    (
        "B-264", "Block B", "General Commission", "M/s Mahendra Kumar Mittal",
        "Mahendra Kumar Mittal", "9811195655", "9810452332",
        "Lic No B-2046", "Wholesale Produce & General Commission Agent", "Various Crates / Sacks",
        "07_General_Produce_Commission/B-264_Mahendra_Kumar_Mittal_Wholesale_Produce.jpg",
        "Old-line mandi commission firm with APMC license B-2046. Excellent for sourcing emergency stock when specific specialty lots run dry.",
        "Store: Reserve stock replenishment. Online: Standard grocery catalog backup."
    ),
    (
        "B-251", "Block B", "General Veg Commission", "M/s Mahesh Saini & Shivam Saini",
        "Mahesh Saini / Shivam Saini", "9811568241", "Mandi Yard",
        "APMC Regd", "General Seasonal Vegetables Wholesale", "Crates (20-25 kg)",
        "07_General_Produce_Commission/B-251_Mahesh_Saini_Shivam_Saini_General_Veg.jpg",
        "Good secondary source for daily mixed vegetables. Fast credit clearance for established buyers.",
        "Store: General display bins. Online: Mixed vegetable packs."
    ),
    (
        "B-287", "Block B", "Wholesale Traders", "M/s Shri Banke Bihari Traders",
        "Shyam Sunder Gupta / Surendra Gupta / Rahul Gupta", "9811157887", "9716599723 / 9911079212",
        "APMC Regd", "Produce Wholesalers & Commission Agents", "Crates & Sacks",
        "07_General_Produce_Commission/B-287_Shri_Banke_Bihari_Traders_Wholesale.jpg",
        "Multi-generation family aadhati firm. Transparent weighment slip (kanta parcha). Check weigh bridge slip.",
        "Store: Standard shelf stock. Online: Regular catalog replenishment."
    ),
    (
        "B-288", "Block B", "Produce Commission", "M/s Brothers Fruit & Vegetable Wholesale",
        "Brothers Firm Representative", "Mandi Desk", "Mandi Yard",
        "Lic No B-240", "Wholesale Fruit & Vegetable Commission", "Crates & Cartons",
        "07_General_Produce_Commission/B-288_Brothers_Fruit_Vegetable_Wholesale.jpg",
        "Registered APMC licensee B-240. Reliable quality grading across mixed fruits and seasonal squash.",
        "Store: General shelf stock. Online: Standard delivery catalog."
    ),
    (
        "B-135", "Block B", "Produce Wholesalers", "M/s Azad Bhai & Company",
        "Irshad Bhai / Rashid Bhai", "9811927040", "9582066551",
        "APMC Regd", "Wholesale Fruits & Vegetables", "Crates (25 kg)",
        "07_General_Produce_Commission/B-135_Azad_Bhai_Co_Produce_Wholesale.jpg",
        "Active trading firm in Block B. Inspect produce directly on arrival. Negotiate on cash payment terms.",
        "Store: Regular retail display. Online: Everyday vegetable packs."
    ),
    (
        "B-101", "Block B", "Wholesale Traders", "M/s R.K. Traders",
        "Ravi Kohli", "9873442731", "Mandi Yard",
        "Lic No 2344", "Fruit & Vegetable Commission Agents", "Crates & Sacks",
        "07_General_Produce_Commission/B-101_RK_Traders_Ravi_Kohli_Wholesale.jpg",
        "APMC license 2344. Experienced aadhati with strong logistics ties to UP & Haryana farmer groups.",
        "Store: General shelf stock. Online: Weekly vegetable subscription backup."
    )
]

for row_idx, vendor in enumerate(vendors_data, 4):
    ws1.row_dimensions[row_idx].height = 28
    is_even = (row_idx % 2 == 0)
    current_fill = fill_zebra if is_even else fill_white
    
    ws1.cell(row=row_idx, column=1, value=row_idx - 3).alignment = align_center
    ws1.cell(row=row_idx, column=2, value=vendor[0]).alignment = align_center # Shop No
    ws1.cell(row=row_idx, column=3, value=vendor[1]).alignment = align_center # Block
    ws1.cell(row=row_idx, column=4, value=vendor[2]).alignment = align_left # Category
    ws1.cell(row=row_idx, column=5, value=vendor[3]).alignment = align_left # Firm Name
    ws1.cell(row=row_idx, column=6, value=vendor[4]).alignment = align_left # Proprietor
    ws1.cell(row=row_idx, column=7, value=vendor[5]).alignment = align_center # Phone 1
    ws1.cell(row=row_idx, column=8, value=vendor[6]).alignment = align_center # Phone 2
    ws1.cell(row=row_idx, column=9, value=vendor[7]).alignment = align_center # License
    ws1.cell(row=row_idx, column=10, value=vendor[8]).alignment = align_left # Produce
    ws1.cell(row=row_idx, column=11, value=vendor[9]).alignment = align_left # Packaging
    ws1.cell(row=row_idx, column=12, value=vendor[10]).alignment = align_left # Photo Ref
    ws1.cell(row=row_idx, column=13, value=vendor[11]).alignment = align_left # How We Buy
    ws1.cell(row=row_idx, column=14, value=vendor[12]).alignment = align_left # How We Sell
    
    for c_idx in range(1, 15):
        cell = ws1.cell(row=row_idx, column=c_idx)
        cell.font = font_regular
        cell.fill = current_fill
        cell.border = border_cell

# Adjust column widths for Sheet 1
col_widths1 = {
    "A": 7, "B": 14, "C": 12, "D": 22, "E": 34, "F": 28, 
    "G": 16, "H": 22, "I": 16, "J": 35, "K": 26, "L": 38, 
    "M": 55, "N": 55
}
for col_letter, width in col_widths1.items():
    ws1.column_dimensions[col_letter].width = width

print("Sheet 1 (Vendor_Directory) populated with 52 entries.")

# -------------------------------------------------------------
# SHEET 2: PROCUREMENT BUYING SOP
# -------------------------------------------------------------
ws2 = wb.create_sheet(title="Procurement_Buying_SOP")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:G1")
ws2["A1"] = "GHAZIPUR SUBZI MANDI — DAILY PROCUREMENT STANDARD OPERATING PROCEDURE (SOP)"
ws2["A1"].font = font_title
ws2["A1"].fill = fill_title
ws2["A1"].alignment = align_center
ws2.row_dimensions[1].height = 28

# Subtitle
ws2.merge_cells("A2:G2")
ws2["A2"] = "Complete guide to wholesale purchasing schedules, APMC market economics, quality grading tests, and transport logistics"
ws2["A2"].font = font_subtitle
ws2["A2"].fill = fill_title
ws2["A2"].alignment = align_center
ws2.row_dimensions[2].height = 20

# Section 1: Daily Schedule
ws2.merge_cells("A4:G4")
ws2["A4"] = "1. DAILY MANDI OPERATIONAL SCHEDULE & TIMELINE"
ws2["A4"].font = font_section
ws2["A4"].fill = fill_section
ws2["A4"].alignment = align_left
ws2.row_dimensions[4].height = 24

sched_headers = ["Time Slot", "Phase", "Activity & Location", "Target Produce", "Key Mandi Shops", "Inspection Priority", "Action Item"]
ws2.row_dimensions[5].height = 25
for c_idx, h in enumerate(sched_headers, 1):
    c = ws2.cell(row=5, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

sched_data = [
    ("03:00 AM - 03:30 AM", "Transit & Dispatch", "Tata Ace / Chota Hathi leaves store for Ghazipur Mandi", "Logistics readiness", "Mandi Gate 1 / Gate 2", "Vehicle cleanliness, empty crates loaded", "Ensure 40 clean plastic crates on vehicle"),
    ("03:30 AM - 04:30 AM", "Auction Phase 1", "Herbs & High-Perishables procurement", "Dhaniya, Pudina, Palak, Methi, Tomatoes", "B-211 (Dhaniya Wale), B-81 (Haji Habib), B-105", "Moisture, rot, yellowing, crate firmness", "Buy 1 ripe crate + 1 semi-ripe crate of tomatoes"),
    ("04:30 AM - 05:15 AM", "Auction Phase 2", "Green Gourds, Exotics & Fruiting Veg", "Bhindi, Lauki, Parwal, Capsicum, Mushroom", "B-242 (Vedram Baba), A-171 (JMD), B-125 (SRTC)", "Tip-snap test on okra; mushroom cap whiteness", "Inspect mushroom packs; ensure cool transit box"),
    ("05:15 AM - 06:15 AM", "Auction Phase 3", "Brassicas, Root Veg & Staples", "Cauliflower, Cabbage, Potatoes, Onions, Garlic", "B-180, B-122, B-238, B-17 (Nashik), B-60", "Head compactness, sprout check, dry onion skin", "Check Sambhal truck tailgate for direct potato lots"),
    ("06:15 AM - 07:00 AM", "Consolidation & Loading", "Vehicle loading, palledari settlement, empty crate balance", "All 40-50 crates & sacks", "Mandi Loading Yard B", "Heavy tubers on floor, delicate greens on top", "Obtain Kanta Parcha (weigh slip) & Aadhati slips"),
    ("07:00 AM - 07:45 AM", "Return Transit & Receiving", "Inbound vehicle returns to store; unloading & misting", "Store intake dock", "Retail Store Front", "Weigh-in verification against mandi parcha", "Sort Grade A to shelves; mist greens; tag prices")
]

for r_idx, row_item in enumerate(sched_data, 6):
    ws2.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws2.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_center if c_idx in [1, 2] else align_left

# Section 2: Financial Structure
ws2.merge_cells("A13:G13")
ws2["A13"] = "2. MANDI FINANCIAL COST BREAKDOWN & TRANSACTION MECHANICS"
ws2["A13"].font = font_section
ws2["A13"].fill = fill_section
ws2["A13"].alignment = align_left
ws2.row_dimensions[13].height = 24

fin_headers = ["Fee / Surcharge Component", "Statutory / Market Rate", "Payable To", "Standard Calculation Basis", "Impact on Cost Price", "Control & Optimization Tactic", "Documentation Required"]
ws2.row_dimensions[14].height = 25
for c_idx, h in enumerate(fin_headers, 1):
    c = ws2.cell(row=14, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

fin_data = [
    ("Mandi APMC Market Fee (Cess)", "1.00% to 1.50%", "Delhi APMC Treasury", "% of gross purchase value on invoice", "+₹0.20 to ₹0.40 / kg", "Ensure purchase through licensed APMC aadhati to avoid unauthorized levying", "APMC Form 6 / Mandi Gate Pass"),
    ("Aadhati Commission (Dami)", "5.00% to 6.00%", "Commission Agent (Aadhati)", "% on gross auction settlement amount", "+₹1.00 to ₹1.80 / kg", "Negotiate to 5% flat based on daily volume commitments across categories", "Aadhati Parcha (Settlement Slip)"),
    ("Palledari (Unloading / Loading)", "₹5 to ₹10 per unit", "Mandi Mathadi / Porters", "Per plastic crate (₹5) / Per 50kg sack (₹8-10)", "+₹0.20 to ₹0.25 / kg", "Have store driver assist to minimize palledari handoffs; fix single loading point", "Daily Palledari cash voucher"),
    ("Crate Security Deposit", "₹50 to ₹100 per crate", "Aadhati / Mandi Yard", "Refundable security deposit per plastic crate", "₹0 net cost (cash flow float)", "Maintain 1-to-1 crate exchange: return yesterday's washed empties to take today's full", "Crate Ledger Entry with Aadhati stamp"),
    ("Vehicle Inbound Logistics", "₹800 to ₹1,200 per trip", "Dedicated Tata Ace Driver", "Daily fixed trip cost (Fuel + Driver allowance)", "+₹1.00 to ₹1.50 / kg (at 800kg load)", "Optimize vehicle fill rate (load 800-1000kg per run) to minimize logistics per kg", "Monthly vehicle logbook & fuel receipt"),
    ("Total Inbound Procurement Load", "8.0% to 11.5%", "Consolidated Mandi Overheads", "Combined effect of statutory, labor & freight fees", "+₹2.40 to ₹3.80 / kg landed markup", "Budget an 8% inbound buffer on wholesale prices to determine accurate landed cost", "Daily Mandi Settlement Sheet")
]

for r_idx, row_item in enumerate(fin_data, 15):
    ws2.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws2.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular if r_idx < 20 else font_bold
        c.fill = fill_accent if r_idx == 20 else (fill_zebra if is_even else fill_white)
        c.border = border_total if r_idx == 20 else border_cell
        c.alignment = align_center if c_idx in [2] else align_left

# Section 3: Quality Inspection SOP
ws2.merge_cells("A22:G22")
ws2["A22"] = "3. COMMODITY QUALITY INSPECTION & DEFECT ACCEPTANCE SOP"
ws2["A22"].font = font_section
ws2["A22"].fill = fill_section
ws2["A22"].alignment = align_left
ws2.row_dimensions[22].height = 24

qual_headers = ["Vegetable Category", "Target Physical Indicators", "Defects to Reject (Zero Tolerance)", "Grade A (Retail & Online)", "Grade B (Processing & Markdown)", "Storage & Transit Care", "Rejection Action"]
ws2.row_dimensions[23].height = 25
for c_idx, h in enumerate(qual_headers, 1):
    c = ws2.cell(row=23, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

qual_data = [
    ("Tomatoes (Desi & Hybrid)", "Firm glossy skin, intact calyx, uniform red/pink color", "Blossom end rot, skin splits, squished fruit, fruit borer holes", "Firm, 70-80% red, uniform 50-60mm diameter", "Over-ripe 95%+ red, minor surface scratches", "Transport in vented crates max 3 layers deep; no direct sunlight", "Demand crate replacement or 20% discount on entire lot"),
    ("Potatoes (Sambhal/Hapur/Agra)", "Dry unwashed skin, dormant eyes, solid weight", "Green sunburn patches (solanine toxin), soft watery rot, sprouting", "Smooth skin, 45-60mm oval/round, free of soil clumps", "Odd shape, small size (<35mm), minor harvester cuts", "Store in shaded, well-ventilated dry space; never cover in airtight plastic", "Refuse sacks with wet bottom or sprouting eyes"),
    ("Onions (Nashik / MP Red)", "Paper-dry outer scales, tight dry necks, deep purple-red", "Thick wet neck (neck rot), sprouting green shoots, black internal mold", "Firm, double-bulb free, 50-65mm diameter, dry crackling skin", "Small size (<40mm), single peel missing, slightly soft neck", "Mesh bags elevated on wooden pallets; air circulation essential", "Randomly slice open 2 onions across 3 sacks to verify core soundness"),
    ("Leafy Greens (Dhaniya, Pudina)", "Vibrant green, crisp turgid stems, fragrant aroma", "Yellowing leaves, slimy rot, heavy mud clumps, wilting", "Roots intact, free of yellow leaves, moisture balanced", "Slight stem wilting (can be revived via hydration soak)", "Wrap in damp jute sheets during transit; keep misted at store", "Reject bundles with rotten black roots or bad fermentation odor"),
    ("Cruciferous (Gobhi & Cabbage)", "Tight compact curds, snow white color, fresh green leaves", "Brown curd spots (Alternaria), aphids, loose curds, yellowing", "Snow-white compact head, 500-800g, protective jacket leaves", "Loose curds, small head (<350g), slight riceyness", "Store head-up in single tiers; do not compress in bags", "Trim damaged florets into pre-cut soup packs"),
    ("Gourds & Okra (Bhindi, Lauki)", "Crisp snapping tip on okra; tender pale skin on lauki", "Fibrous tough okra tips, yellowing gourds, insect stings", "Bright green 2-3 inch okra, zero stringiness; pale green tender lauki", "Over-matured tough pods, curved gourds, superficial scars", "Keep separated in ventilated crates; mist lightly every 2 hours", "Test 5 okra pods per basket; if >1 fails snap test, negotiate B-grade price"),
    ("Exotics (Mushrooms & Corn)", "Closed button caps, clean white veil, dry surface", "Opened gills showing black spores, slimy brown patches, sour odor", "Clean white caps, 25-35mm diameter, tight veil", "Slight cap opening, minor bruising (use for cut mushroom)", "Maintain 4°C - 8°C cold chain; insulated cooler box in vehicle", "Reject punnets with condensation inside or dark brown caps")
]

for r_idx, row_item in enumerate(qual_data, 24):
    ws2.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws2.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_left

col_widths2 = {"A": 22, "B": 24, "C": 35, "D": 28, "E": 28, "F": 35, "G": 32}
for col_letter, width in col_widths2.items():
    ws2.column_dimensions[col_letter].width = width

print("Sheet 2 (Procurement_Buying_SOP) generated.")

# -------------------------------------------------------------
# SHEET 3: OFFLINE STORE OPERATIONS
# -------------------------------------------------------------
ws3 = wb.create_sheet(title="Offline_Store_Operations")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells("A1:G1")
ws3["A1"] = "OFFLINE RETAIL STORE — OPERATIONAL BLUEPRINT & ZERO-WASTAGE PROTOCOL"
ws3["A1"].font = font_title
ws3["A1"].fill = fill_title
ws3["A1"].alignment = align_center
ws3.row_dimensions[1].height = 28

ws3.merge_cells("A2:G2")
ws3["A2"] = "Physical store merchandising, hourly operational rhythm, shrinkage mitigation, and Point-of-Sale (POS) weighing protocols"
ws3["A2"].font = font_subtitle
ws3["A2"].fill = fill_title
ws3["A2"].alignment = align_center
ws3.row_dimensions[2].height = 20

# Section 1: Store Blueprint
ws3.merge_cells("A4:G4")
ws3["A4"] = "1. STORE ZONING & MERCHANDISING ARCHITECTURE (350 - 500 SQ FT MODEL)"
ws3["A4"].font = font_section
ws3["A4"].fill = fill_section
ws3["A4"].alignment = align_left
ws3.row_dimensions[4].height = 24

zone_headers = ["Store Zone", "Location in Store", "Target Commodities", "Fixture & Equipment Type", "Ideal Temp & Humidity", "Merchandising & Display Technique", "Customer Experience Goal"]
ws3.row_dimensions[5].height = 25
for c_idx, h in enumerate(zone_headers, 1):
    c = ws3.cell(row=5, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

zone_data = [
    ("Entrance Promo Island", "Front Center (Visible from street)", "Seasonal Fruits, Watermelons, Mangoes, Winter Peas", "Rustic wooden feature tiers / barrels", "Ambient (24°C - 28°C)", "Mass visual abundance display; high-aroma seasonal draw", "Stop passerby foot traffic; create fresh farm perception"),
    ("Hydration Green Wall", "Left Perimeter Wall (Eye Level)", "Dhaniya, Pudina, Palak, Methi, Exotic Salads, Herbs", "3-Tier inclined SS rack with ultrasonic micro-misters", "18°C - 22°C | 85-90% RH", "Lined with waterproof trays; fine misting every 45 mins", "Visual crispness; eliminate leaf wilting; premium organic look"),
    ("Gondola Produce Aisle", "Center Floor (Two-sided access)", "Tomatoes, Capsicum, Bhindi, Gourds, Cucumbers, Beans", "Modular stackable wooden crates with green astro-turf", "22°C - 25°C | 70% RH", "Color-contrast blocking (Red tomato next to green capsicum)", "Easy browsing; spacious aisles for basket carts"),
    ("Allium & Tuber Bunkers", "Right Perimeter Wall", "Potatoes (Chipsona, Desi), Onions (Nashik), Garlic, Arbi", "Sloped wooden bins with bottom ventilation slats", "Ambient Dry (22°C - 26°C)", "Bulk gravity feed; bottom clearance for dirt separation", "High velocity staple pick; easy self-selection into paper bags"),
    ("Commercial Chiller", "Back Wall (Adjacent to Prep Room)", "Mushrooms, Paneer, Sweet Corn, Sprouts, Cut Veg, Tofu", "Double-door glass display commercial refrigerator", "4°C - 8°C | Controlled", "Organized by meal categories (Continental / North Indian)", "Hygiene confidence; cross-selling high margin dairy/exotics"),
    ("Weighing & POS Checkout", "Front Right (Exit side)", "Impulse add-ons (Lemons, Green Chillies, Ginger, Garlic)", "Digital barcode scale + POS touch terminal + UPI QR", "Ambient air conditioned", "Scale auto-transmits weight to POS; fast thermal billing", "Zero checkout queues; average transaction time < 45 seconds")
]

for r_idx, row_item in enumerate(zone_data, 6):
    ws3.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws3.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_left

# Section 2: Daily Rhythm
ws3.merge_cells("A13:G13")
ws3["A13"] = "2. DAILY HOURLY OPERATIONAL RHYTHM"
ws3["A13"].font = font_section
ws3["A13"].fill = fill_section
ws3["A13"].alignment = align_left
ws3.row_dimensions[13].height = 24

rhythm_headers = ["Time Window", "Operational Phase", "Core Activities", "Primary Responsible Staff", "Wastage Control Check", "POS / System Action", "Customer Facing State"]
ws3.row_dimensions[14].height = 25
for c_idx, h in enumerate(rhythm_headers, 1):
    c = ws3.cell(row=14, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

rhythm_data = [
    ("07:00 AM - 08:00 AM", "Intake & Shelving", "Unload vehicle, check mandi parcha weights, trim stem ends, mist leafy greens, stock displays", "Store Supervisor & Associate", "Record inbound intake weight vs Mandi parcha; flag shortages", "Update daily purchase prices in POS; print new price tags", "Closed for setup; fresh aroma buildup"),
    ("08:00 AM - 12:30 PM", "Peak Morning Rush", "Assist homemakers, continuous weighing, maintain shelf fullness, offer free dhaniya/mirchi", "Cashier & 2 Sales Associates", "Hourly shelf culling of crushed or bruised items", "Process cash & UPI billing; capture customer mobile for WhatsApp", "Full vibrancy; high-energy service; zero queues"),
    ("12:30 PM - 04:00 PM", "Midday Maintenance & Online Prep", "Restock bins from reserve; second misting cycle; grade items; pack online delivery pre-orders", "Store Supervisor & Packing Staff", "Trim outer leaves of cauliflowers/greens; move Grade B to prep room", "Inventory cycle count on top 10 SKUs; sync stock with WhatsApp catalog", "Quiet browsing; personalized service for elders"),
    ("04:00 PM - 07:30 PM", "Peak Evening Rush", "Cater to working professionals & dinner shoppers; promote ready-to-cook packs & salad kits", "Full Store Team (3 Staff)", "Keep displays neat; re-stack potato/onion bins", "Fast checkout; promote tomorrow morning delivery slots", "High footfall; active sampling of fruits/salads"),
    ("07:30 PM - 09:15 PM", "Happy Hour Clearance", "Activate 'Happy Hour' 25-30% discount on delicate greens, ripe tomatoes & cut packs to clear stock", "Sales Associates", "Execute markdown pricing; separate items unfit for next day", "Apply markdown barcode in POS system", "'Fresh Clearance' signage; price-sensitive buyers"),
    ("09:15 PM - 10:00 PM", "Store Close & EOD Reconciliation", "Store shutdown, deep sweeping & washdown, misting equipment purge, inventory count, cash reconciliation", "Store Supervisor", "Total daily shrinkage weighed & recorded in ERP (< 4.5% target)", "Close POS day batch; transmit morning procurement indent to buyer", "Store closed; night ventilation active")
]

for r_idx, row_item in enumerate(rhythm_data, 15):
    ws3.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws3.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_left

# Section 3: 4-Tier Wastage Control
ws3.merge_cells("A22:G22")
ws3["A22"] = "3. 4-TIER ZERO-WASTAGE & SHRINKAGE MITIGATION PROTOCOL"
ws3["A22"].font = font_section
ws3["A22"].fill = fill_section
ws3["A22"].alignment = align_left
ws3.row_dimensions[22].height = 24

waste_headers = ["Tier Level", "Produce State / Condition", "Handling Process & Action", "Target Margin / Pricing", "Target Consumer / Channel", "Target SKUs", "Expected Recovery %"]
ws3.row_dimensions[23].height = 25
for c_idx, h in enumerate(waste_headers, 1):
    c = ws3.cell(row=23, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

waste_data = [
    ("Tier 1: Fresh Premium", "Pristine, Grade A, unblemished produce", "Washed, misted, displayed on main shelves & online store", "Full Markup (35% - 45% Gross Margin)", "Walk-in retail shoppers & online app customers", "All fresh arrivals (Tomatoes, Greens, Tubers)", "75% to 80% of total stock"),
    ("Tier 2: Ready-to-Cook Processing", "Cosmetically imperfect, odd-sized, or superficial blemishes", "Peeled, trimmed, sanitized, diced & vacuum-sealed in prep room", "High Value-Add Markup (40% - 60% Gross Margin)", "Busy working couples, bachelor meal preps, cafes", "Peeled Garlic, Diced Pumpkin, Cut Kathal, Sambar Veg Kit", "10% to 12% of total stock"),
    ("Tier 3: Evening Markdown", "Delicate greens & ripe tomatoes remaining at 7:30 PM", "Bunched into clearance baskets with 25% - 30% discount sticker", "Break-Even Price (0% to +10% Gross Margin)", "Budget-conscious neighbourhood shoppers & domestic cooks", "Ripe Tomatoes, Spinach, Coriander, Cucumbers", "6% to 8% of total stock"),
    ("Tier 4: B2B Liquidate & Organic Compost", "Over-ripe tomatoes, softening produce, trimmings/peels", "Liquidate to local street food carts / cloud kitchens; peelings to organic compost", "At Cost (₹0 loss) / Peels to urban gardening compost kits", "Local roadside food stalls, juice centres, urban gardeners", "Soft Tomatoes (for puree), Overripe bananas, vegetable peels", "Remaining 2% to 3% (Zero landfill waste)")
]

for r_idx, row_item in enumerate(waste_data, 24):
    ws3.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws3.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_left

col_widths3 = {"A": 22, "B": 24, "C": 35, "D": 26, "E": 28, "F": 35, "G": 32}
for col_letter, width in col_widths3.items():
    ws3.column_dimensions[col_letter].width = width

print("Sheet 3 (Offline_Store_Operations) generated.")

# -------------------------------------------------------------
# SHEET 4: ONLINE DELIVERY MODEL
# -------------------------------------------------------------
ws4 = wb.create_sheet(title="Online_Delivery_Model")
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells("A1:G1")
ws4["A1"] = "ONLINE VEGETABLE DELIVERY — HYPERLOCAL E-COMMERCE ROLLOUT MODEL"
ws4["A1"].font = font_title
ws4["A1"].fill = fill_title
ws4["A1"].alignment = align_center
ws4.row_dimensions[1].height = 28

ws4.merge_cells("A2:G2")
ws4["A2"] = "Phase 2 expansion: Store-as-a-Hub omnichannel architecture, WhatsApp Business catalog, slot vs express delivery, and packaging SOP"
ws4["A2"].font = font_subtitle
ws4["A2"].fill = fill_title
ws4["A2"].alignment = align_center
ws4.row_dimensions[2].height = 20

# Section 1: Ordering & Tech Architecture
ws4.merge_cells("A4:G4")
ws4["A4"] = "1. ORDERING CHANNELS & TECH STACK"
ws4["A4"].font = font_section
ws4["A4"].fill = fill_section
ws4["A4"].alignment = align_left
ws4.row_dimensions[4].height = 24

tech_headers = ["Channel", "Platform / Tool", "Customer User Journey", "Payment Integration", "Inventory Sync", "Target Demographic", "Operational Advantage"]
ws4.row_dimensions[5].height = 25
for c_idx, h in enumerate(tech_headers, 1):
    c = ws4.cell(row=5, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

tech_data = [
    ("WhatsApp Business Catalog", "WhatsApp Business API / Interakt / Wati", "Customer sends 'Hi' -> Browses photo catalog -> Adds to cart -> One-click order", "UPI Deep Link / Dynamic QR / Cash on Delivery", "Updated 2x daily (8 AM & 2 PM) based on POS shelf count", "Homemakers, apartment society residents, seniors", "Zero app download friction; 98% open rates; instant push notifications"),
    ("Hyperlocal Web Store", "PWA (Shopify / Dukaan / Custom Web App)", "Customer opens link from WhatsApp/Google -> Selects delivery slot -> Enters address -> Pays", "Razorpay / PhonePe Payment Gateway / Pay on Delivery", "Real-time 2-way webhook sync with store POS barcode scale", "Working professionals, tech-savvy millennials, recurring buyers", "Supports scheduled subscriptions; automated cart abandonment reminders"),
    ("Apartment Society Kiosk / Group Buying", "WhatsApp Community Groups (Tower-wise)", "Daily morning broadcast of 'Fresh Harvest of the Day' -> Pre-orders bundled by tower", "Consolidated UPI payment / Monthly household ledger", "Consolidated picking list generated at 7:00 AM", "Gated high-rise residential societies within 3 km", "Zero customer acquisition cost; massive drop density (10-15 orders per tower)")
]

for r_idx, row_item in enumerate(tech_data, 6):
    ws4.row_dimensions[r_idx].height = 26
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws4.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_left

# Section 2: Delivery Windows & Fleet Logistics
ws4.merge_cells("A10:G10")
ws4["A10"] = "2. DELIVERY RADIUS, WINDOWS & FLEET LOGISTICS"
ws4["A10"].font = font_section
ws4["A10"].fill = fill_section
ws4["A10"].alignment = align_left
ws4.row_dimensions[10].height = 24

fleet_headers = ["Delivery Mode", "Service Radius", "Order Cutoff Time", "Delivery Window", "Vehicle Type", "Fulfillment Cost / Order", "Customer Delivery Fee"]
ws4.row_dimensions[11].height = 25
for c_idx, h in enumerate(fleet_headers, 1):
    c = ws4.cell(row=11, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

fleet_data = [
    ("Morning Scheduled Slot", "0 to 3.5 km", "Orders placed before 10:00 PM previous night or 6:30 AM", "07:30 AM - 09:30 AM", "Electric 2-Wheeler (EV Scooter) with delivery bag", "₹22 to ₹26 per order (batched 6-8 drops/run)", "FREE on orders >= ₹199 (₹25 if below ₹199)"),
    ("Midday Scheduled Slot", "0 to 3.5 km", "Orders placed before 11:30 AM", "12:00 PM - 02:00 PM", "Electric 2-Wheeler", "₹25 to ₹28 per order", "FREE on orders >= ₹199 (₹25 if below ₹199)"),
    ("Evening Scheduled Slot", "0 to 4.5 km", "Orders placed before 04:30 PM", "05:30 PM - 07:30 PM", "Electric 2-Wheeler", "₹24 to ₹28 per order", "FREE on orders >= ₹199 (₹25 if below ₹199)"),
    ("Hyperlocal Express (60 Min)", "0 to 2.5 km", "Orders placed between 09:00 AM and 07:00 PM", "Within 45 to 60 Minutes", "Store dedicated runner / Porter On-Demand backup", "₹35 to ₹42 per order (single point delivery)", "Flat ₹35 Express Convenience Fee")
]

for r_idx, row_item in enumerate(fleet_data, 12):
    ws4.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws4.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_center if c_idx in [2, 7] else align_left

# Section 3: Packaging Standards
ws4.merge_cells("A17:G17")
ws4["A17"] = "3. PACKAGING SPECIFICATIONS & SUSTAINABLE PACKING SOP"
ws4["A17"].font = font_section
ws4["A17"].fill = fill_section
ws4["A17"].alignment = align_left
ws4.row_dimensions[17].height = 24

pack_headers = ["Vegetable Category", "Primary Packaging Material", "Secondary / Outer Layer", "Ventilation / Breathability", "Cost per Package", "Customer Unboxing Experience", "Freshness Retention Period"]
ws4.row_dimensions[18].height = 25
for c_idx, h in enumerate(pack_headers, 1):
    c = ws4.cell(row=18, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

pack_data = [
    ("Potatoes & Onions (1kg / 2kg / 5kg)", "Breathable unbleached brown Kraft paper sack / Cotton net", "Recyclable kraft paper carry bag", "Micro-perforated holes for air circulation", "₹2.20 to ₹3.50", "Clean, rustic, zero soil spillage; reusable bag", "10 to 14 days at ambient temp"),
    ("Tomatoes & Fruiting Veg (500g / 1kg)", "Ventilated rigid RPET punnet / Padded paper pouch", "Separated from heavy tubers", "4 corner punch ventilation slots", "₹2.80 to ₹3.80", "Flawless, zero bruising, uniform red aesthetics", "5 to 7 days in refrigerator"),
    ("Fresh Herbs (Dhaniya, Pudina 100g)", "Food-grade moisture-absorbent parchment paper wrap", "Sealed outer kraft paper pouch", "Moisture-lock wrap keeps stems crisp", "₹1.20 to ₹1.60", "Farm-fresh fragrance upon tearing wrap open", "4 to 6 days in refrigerator"),
    ("Okra, Beans & Gourds (500g)", "Gusseted Kraft paper bag with clear glassine window", "Main grocery tote", "Full side air permeability", "₹2.00 to ₹2.50", "Clear window displays pod freshness before opening", "4 to 5 days"),
    ("Mushrooms & Exotics (200g)", "Molded pulp tray with perforated breathable cling wrap", "Insulated cold pouch with gel ice pad", "Cold chain maintained at 4°C - 8°C", "₹3.50 to ₹4.50", "Gourmet supermarket finish; pristine white caps", "6 to 8 days in cold storage")
]

for r_idx, row_item in enumerate(pack_data, 19):
    ws4.row_dimensions[r_idx].height = 24
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws4.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_center if c_idx in [5] else align_left

col_widths4 = {"A": 24, "B": 28, "C": 35, "D": 26, "E": 22, "F": 35, "G": 28}
for col_letter, width in col_widths4.items():
    ws4.column_dimensions[col_letter].width = width

print("Sheet 4 (Online_Delivery_Model) generated.")

# -------------------------------------------------------------
# SHEET 5: UNIT ECONOMICS & PRICING
# -------------------------------------------------------------
ws5 = wb.create_sheet(title="Unit_Economics_&_Pricing")
ws5.views.sheetView[0].showGridLines = True

ws5.merge_cells("A1:Q1")
ws5["A1"] = "OMNICHANNEL VEGETABLE BUSINESS — SKU UNIT ECONOMICS & MARGIN CALCULATOR"
ws5["A1"].font = font_title
ws5["A1"].fill = fill_title
ws5["A1"].alignment = align_center
ws5.row_dimensions[1].height = 28

ws5.merge_cells("A2:Q2")
ws5["A2"] = "Dynamic pricing engine comparing Mandi Landed Costs, Shrinkage Allowances, Offline Retail Margins, and Online E-Commerce Profitability across 25 Top SKUs"
ws5["A2"].font = font_subtitle
ws5["A2"].fill = fill_title
ws5["A2"].alignment = align_center
ws5.row_dimensions[2].height = 20

headers5 = [
    "S.No", "Vegetable / SKU Name", "Category", "Mandi Base Price (₹/kg)",
    "Inbound Mandi Overheads (8%) (₹/kg)", "Landed Cost (₹/kg)",
    "Shrinkage / Wastage Allowance (%)", "Net Effective Cost (₹/kg)",
    "Offline Retail Price (₹/kg)", "Offline Gross Margin (₹/kg)", "Offline Margin %",
    "Online Pack Size (kg)", "Online Packaging Material (₹/pack)", "Online Selling Price (₹/pack)",
    "Online Gross Margin (₹/pack)", "Online Margin %", "Projected Daily Volume (kg)"
]

ws5.row_dimensions[3].height = 32
for c_idx, h in enumerate(headers5, 1):
    c = ws5.cell(row=3, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

# 25 High-Volume SKUs
sku_data = [
    # (Name, Category, Mandi Base, Shrinkage %, Offline Retail Price, Online Pack Size, Online Packaging Cost, Online Price, Daily Vol)
    ("Potato - Regular (Sambhal/Agra)", "Potato & Onion", 16.00, 0.03, 26.00, 1.0, 2.50, 29.00, 90),
    ("Potato - Chipsona / Surya (Sugar Free)", "Potato & Onion", 22.00, 0.03, 35.00, 1.0, 3.00, 39.00, 45),
    ("Potato - Hapur Desi Special", "Potato & Onion", 20.00, 0.03, 32.00, 1.0, 2.50, 36.00, 35),
    ("Onion - Nashik Red Premium", "Potato & Onion", 28.00, 0.04, 44.00, 1.0, 2.50, 49.00, 80),
    ("Onion - Regular MP Yellow/Red", "Potato & Onion", 22.00, 0.04, 34.00, 1.0, 2.50, 38.00, 50),
    ("Garlic - Desi Large Clove", "Potato & Onion", 120.00, 0.02, 180.00, 0.25, 2.00, 55.00, 15),
    ("Ginger (Adrak) - Bangalore Washed", "Potato & Onion", 70.00, 0.04, 110.00, 0.25, 1.80, 35.00, 18),
    ("Tomato - Hybrid Firm (Shivaji)", "Tomato Specialists", 26.00, 0.06, 42.00, 1.0, 3.00, 48.00, 75),
    ("Tomato - Desi Tangy Salad", "Tomato Specialists", 22.00, 0.07, 36.00, 1.0, 3.00, 42.00, 55),
    ("Dhaniya (Fresh Coriander)", "Green Vegetables & Herbs", 35.00, 0.12, 65.00, 0.10, 1.50, 12.00, 20),
    ("Pudina (Fresh Mint Leaves)", "Green Vegetables & Herbs", 40.00, 0.10, 75.00, 0.10, 1.50, 14.00, 10),
    ("Spinach (Palak - Washed)", "Green Vegetables & Herbs", 20.00, 0.12, 38.00, 0.50, 2.00, 25.00, 25),
    ("Bhindi (Okra / Lady Finger)", "Green Vegetables & Herbs", 34.00, 0.06, 55.00, 0.50, 2.20, 34.00, 35),
    ("Cauliflower (Phool Gobhi)", "Green Vegetables & Herbs", 25.00, 0.08, 44.00, 0.60, 2.50, 35.00, 40),
    ("Cabbage (Patta Gobhi)", "Green Vegetables & Herbs", 18.00, 0.06, 30.00, 0.70, 2.00, 28.00, 35),
    ("Capsicum - Green (Shimla Mirch)", "Green Vegetables & Herbs", 42.00, 0.05, 68.00, 0.50, 2.50, 42.00, 25),
    ("Bottle Gourd (Lauki - Tender)", "Green Vegetables & Herbs", 20.00, 0.05, 34.00, 1.0, 2.00, 39.00, 35),
    ("Pointed Gourd (Parwal)", "Green Vegetables & Herbs", 45.00, 0.05, 75.00, 0.50, 2.20, 45.00, 15),
    ("Cucumber (Desi Kheera)", "Green Vegetables & Herbs", 24.00, 0.05, 40.00, 0.50, 2.00, 26.00, 30),
    ("Green Chilli (Hari Mirch)", "Green Vegetables & Herbs", 55.00, 0.04, 90.00, 0.20, 1.50, 24.00, 15),
    ("Radish (Mooli with Leaves)", "Root Vegetables", 18.00, 0.08, 30.00, 1.0, 2.50, 35.00, 25),
    ("Carrot (Gajar - Seasonal)", "Root Vegetables", 28.00, 0.05, 45.00, 1.0, 2.50, 52.00, 30),
    ("Button Mushroom (200g Punnet)", "Exotics & Frozen", 32.00, 0.03, 55.00, 0.20, 3.50, 60.00, 25), # Note: Mandi base per punnet
    ("Sweet Corn (Pack of 2 Cobs)", "Exotics & Frozen", 24.00, 0.03, 40.00, 0.40, 2.50, 45.00, 20),
    ("Bananas - Robusta G9 (Dozen)", "Fruits & Coconut", 32.00, 0.05, 52.00, 1.2, 3.00, 58.00, 40)
]

for idx, sku in enumerate(sku_data, 4):
    row_num = idx
    ws5.row_dimensions[row_num].height = 24
    is_even = (row_num % 2 == 0)
    current_fill = fill_zebra if is_even else fill_white
    
    # S.No
    ws5.cell(row=row_num, column=1, value=idx - 3).alignment = align_center
    # Name
    ws5.cell(row=row_num, column=2, value=sku[0]).alignment = align_left
    # Category
    ws5.cell(row=row_num, column=3, value=sku[1]).alignment = align_left
    # Mandi Base Price
    c_mandi = ws5.cell(row=row_num, column=4, value=sku[2])
    c_mandi.alignment = align_right
    c_mandi.number_format = '₹#,##0.00'
    # Mandi Overheads (8% formula)
    c_overhead = ws5.cell(row=row_num, column=5, value=f"=D{row_num}*0.08")
    c_overhead.alignment = align_right
    c_overhead.number_format = '₹#,##0.00'
    # Landed Cost (formula)
    c_landed = ws5.cell(row=row_num, column=6, value=f"=D{row_num}+E{row_num}")
    c_landed.alignment = align_right
    c_landed.number_format = '₹#,##0.00'
    # Shrinkage %
    c_shrink = ws5.cell(row=row_num, column=7, value=sku[3])
    c_shrink.alignment = align_center
    c_shrink.number_format = '0.0%'
    # Net Effective Cost (formula: Landed / (1 - Shrinkage%))
    c_net = ws5.cell(row=row_num, column=8, value=f"=F{row_num}/(1-G{row_num})")
    c_net.alignment = align_right
    c_net.number_format = '₹#,##0.00'
    # Offline Retail Price
    c_off_price = ws5.cell(row=row_num, column=9, value=sku[4])
    c_off_price.alignment = align_right
    c_off_price.number_format = '₹#,##0.00'
    # Offline Gross Margin ₹ (formula: Retail - Net Cost)
    c_off_mgn = ws5.cell(row=row_num, column=10, value=f"=I{row_num}-H{row_num}")
    c_off_mgn.alignment = align_right
    c_off_mgn.number_format = '₹#,##0.00'
    # Offline Margin % (formula: Margin ₹ / Retail Price)
    c_off_pct = ws5.cell(row=row_num, column=11, value=f"=J{row_num}/I{row_num}")
    c_off_pct.alignment = align_right
    c_off_pct.number_format = '0.0%'
    # Online Pack Size
    c_pack_size = ws5.cell(row=row_num, column=12, value=sku[5])
    c_pack_size.alignment = align_center
    c_pack_size.number_format = '0.00'
    # Online Packaging Material
    c_pkg_cost = ws5.cell(row=row_num, column=13, value=sku[6])
    c_pkg_cost.alignment = align_right
    c_pkg_cost.number_format = '₹#,##0.00'
    # Online Selling Price
    c_on_price = ws5.cell(row=row_num, column=14, value=sku[7])
    c_on_price.alignment = align_right
    c_on_price.number_format = '₹#,##0.00'
    # Online Gross Margin ₹ (formula: Online Price - (Net Cost * Pack Size) - Packaging Cost)
    c_on_mgn = ws5.cell(row=row_num, column=15, value=f"=N{row_num}-(H{row_num}*L{row_num})-M{row_num}")
    c_on_mgn.alignment = align_right
    c_on_mgn.number_format = '₹#,##0.00'
    # Online Margin % (formula: Online Margin ₹ / Online Selling Price)
    c_on_pct = ws5.cell(row=row_num, column=16, value=f"=O{row_num}/N{row_num}")
    c_on_pct.alignment = align_right
    c_on_pct.number_format = '0.0%'
    # Projected Daily Volume
    c_vol = ws5.cell(row=row_num, column=17, value=sku[8])
    c_vol.alignment = align_right
    c_vol.number_format = '#,##0'
    
    for c_i in range(1, 18):
        cell = ws5.cell(row=row_num, column=c_i)
        cell.font = font_regular
        cell.fill = current_fill
        cell.border = border_cell

# Totals & Summary KPI Block
tot_row = 29
ws5.row_dimensions[tot_row].height = 26
ws5.merge_cells(f"A{tot_row}:C{tot_row}")
ws5[f"A{tot_row}"] = "TOTALS & PORTFOLIO AVERAGES"
ws5[f"A{tot_row}"].font = font_bold
ws5[f"A{tot_row}"].alignment = align_center
ws5[f"A{tot_row}"].fill = fill_subhdr

for c_idx in range(1, 18):
    c = ws5.cell(row=tot_row, column=c_idx)
    c.border = border_total
    c.fill = fill_subhdr

ws5[f"D{tot_row}"] = "=AVERAGE(D4:D28)"
ws5[f"D{tot_row}"].number_format = '₹#,##0.00'
ws5[f"D{tot_row}"].font = font_bold
ws5[f"D{tot_row}"].alignment = align_right

ws5[f"F{tot_row}"] = "=AVERAGE(F4:F28)"
ws5[f"F{tot_row}"].number_format = '₹#,##0.00'
ws5[f"F{tot_row}"].font = font_bold
ws5[f"F{tot_row}"].alignment = align_right

ws5[f"H{tot_row}"] = "=AVERAGE(H4:H28)"
ws5[f"H{tot_row}"].number_format = '₹#,##0.00'
ws5[f"H{tot_row}"].font = font_bold
ws5[f"H{tot_row}"].alignment = align_right

ws5[f"I{tot_row}"] = "=AVERAGE(I4:I28)"
ws5[f"I{tot_row}"].number_format = '₹#,##0.00'
ws5[f"I{tot_row}"].font = font_bold
ws5[f"I{tot_row}"].alignment = align_right

ws5[f"J{tot_row}"] = "=AVERAGE(J4:J28)"
ws5[f"J{tot_row}"].number_format = '₹#,##0.00'
ws5[f"J{tot_row}"].font = font_bold
ws5[f"J{tot_row}"].alignment = align_right

ws5[f"K{tot_row}"] = "=AVERAGE(K4:K28)"
ws5[f"K{tot_row}"].number_format = '0.0%'
ws5[f"K{tot_row}"].font = font_bold
ws5[f"K{tot_row}"].alignment = align_right

ws5[f"P{tot_row}"] = "=AVERAGE(P4:P28)"
ws5[f"P{tot_row}"].number_format = '0.0%'
ws5[f"P{tot_row}"].font = font_bold
ws5[f"P{tot_row}"].alignment = align_right

ws5[f"Q{tot_row}"] = "=SUM(Q4:Q28)"
ws5[f"Q{tot_row}"].number_format = '#,##0 kg'
ws5[f"Q{tot_row}"].font = font_bold
ws5[f"Q{tot_row}"].alignment = align_right

# Monthly Profit & Loss Pro-Forma Block below
ws5.merge_cells("B32:H32")
ws5["B32"] = "MONTHLY STORE FINANCIAL PROJECTIONS (STABILIZED MONTH 3)"
ws5["B32"].font = font_section
ws5["B32"].fill = fill_section
ws5["B32"].alignment = align_left
ws5.row_dimensions[32].height = 24

pnl_items = [
    ("Daily Sales Volume (Combined Offline + Online)", "850 kg / day", "70% Offline Store / 30% Online Delivery"),
    ("Average Gross Realization Price per kg", "₹48.50 / kg", "Blended across staples, greens & exotics"),
    ("Gross Daily Revenue", "₹41,225 / day", "Monthly Revenue (30 days): ₹12,36,750"),
    ("Blended Gross Margin %", "37.5%", "Calculated after all mandi shrinkage & overheads"),
    ("Monthly Gross Profit", "₹4,63,781 / month", "Gross cash generated before operating expenses"),
    ("Operating Expense: Retail Store Rent (400 sq ft)", "₹45,000 / month", "High footfall residential catchment market"),
    ("Operating Expense: Store Staff (1 Sup + 2 Sales + 1 Pack)", "₹62,000 / month", "4 full-time staff salaries + incentives"),
    ("Operating Expense: Delivery Runners (2 EV Scooters)", "₹32,000 / month", "2 delivery staff + fuel/battery charging"),
    ("Operating Expense: Electricity & Cold Storage AC", "₹18,000 / month", "Chillers, misters, store AC & lighting"),
    ("Operating Expense: Daily Mandi Vehicle (Tata Ace)", "₹30,000 / month", "₹1,000 per daily mandi transit run"),
    ("Operating Expense: Packaging Materials & Consumables", "₹22,000 / month", "Kraft bags, paper pouches, punnets, barcode rolls"),
    ("Operating Expense: Marketing, WhatsApp API & Software", "₹12,000 / month", "Flyers, WhatsApp API messages, POS subscription"),
    ("Total Monthly Operating Expenses", "₹2,21,000 / month", "Fixed & operational store overheads"),
    ("NET MONTHLY OPERATING PROFIT (EBITDA)", "₹2,42,781 / month", "19.6% Net Profit Margin on Gross Revenue!"),
    ("Projected Annual Net Earnings", "₹29,13,375 / year", "Substantial annual return on investment")
]

for p_idx, p_item in enumerate(pnl_items, 33):
    ws5.row_dimensions[p_idx].height = 22
    is_tot = "NET MONTHLY" in p_item[0] or "Total Monthly" in p_item[0]
    ws5.cell(row=p_idx, column=2, value=p_item[0]).alignment = align_left
    ws5.cell(row=p_idx, column=3, value=p_item[1]).alignment = align_right
    ws5.merge_cells(f"D{p_idx}:H{p_idx}")
    ws5.cell(row=p_idx, column=4, value=p_item[2]).alignment = align_left
    
    for c_i in range(2, 9):
        cell = ws5.cell(row=p_idx, column=c_i)
        cell.font = font_bold if is_tot else font_regular
        cell.fill = fill_accent if "NET" in p_item[0] else (fill_zebra if p_idx % 2 == 0 else fill_white)
        cell.border = border_total if is_tot else border_cell

col_widths5 = {
    "A": 7, "B": 36, "C": 24, "D": 22, "E": 24, "F": 20, 
    "G": 18, "H": 22, "I": 22, "J": 22, "K": 16, "L": 18, 
    "M": 22, "N": 22, "O": 22, "P": 16, "Q": 20
}
for col_letter, width in col_widths5.items():
    ws5.column_dimensions[col_letter].width = width

print("Sheet 5 (Unit_Economics_&_Pricing) generated.")

# -------------------------------------------------------------
# SHEET 6: PHASED ROADMAP 30-60-90
# -------------------------------------------------------------
ws6 = wb.create_sheet(title="Phased_Roadmap_30_60_90")
ws6.views.sheetView[0].showGridLines = True

ws6.merge_cells("A1:G1")
ws6["A1"] = "30-60-90 DAY OMNICHANNEL ROLLOUT MASTER PLAN"
ws6["A1"].font = font_title
ws6["A1"].fill = fill_title
ws6["A1"].alignment = align_center
ws6.row_dimensions[1].height = 28

ws6.merge_cells("A2:G2")
ws6["A2"] = "Step-by-step phased execution roadmap from initial mandi onboarding to offline store launch, WhatsApp delivery, and e-commerce scaling"
ws6["A2"].font = font_subtitle
ws6["A2"].fill = fill_title
ws6["A2"].alignment = align_center
ws6.row_dimensions[2].height = 20

roadmap_headers = ["Phase & Period", "Strategic Milestone", "Key Actions & Operations", "Target Mandi Vendors / Partners", "Deliverables & Output", "Budget Allocation (₹)", "Target KPI / Metric"]
ws6.row_dimensions[4].height = 28
for c_idx, h in enumerate(roadmap_headers, 1):
    c = ws6.cell(row=4, column=c_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_header
    c.border = border_header

roadmap_data = [
    (
        "Phase 1: Days 1 - 10", "Mandi Relationships & Sourcing Setup",
        "1. Visit Ghazipur Mandi at 4:00 AM across 3 consecutive days.\n2. Meet key Aadhati contacts in Potato/Onion (B-238, B-17, B-60), Tomatoes (B-81, B-105), Herbs (B-211), and Exotics (A-171).\n3. Establish trade credit / weekly ledger accounts.\n4. Secure dedicated Tata Ace logistics driver on monthly retainer.",
        "B-238 (Deepak), B-81 (Haji Habib), B-211 (Naresh Kumar), B-17 (Shyam Sunder)",
        "Signed Aadhati accounts, established crate exchange protocol, fixed daily vehicle route",
        "₹50,000 (Mandi deposits + initial working capital float)",
        "Zero logistics delays; 100% SKU availability at opening"
    ),
    (
        "Phase 1: Days 11 - 20", "Offline Retail Lease & Store Fit-Out",
        "1. Finalize commercial lease for 350-500 sq ft retail space in dense residential catchment.\n2. Construct rustic sloped wooden bins for potatoes/onions.\n3. Install SS hydration misting rack for leafy greens.\n4. Install commercial 2-door glass refrigerator for mushrooms/dairy.\n5. Setup POS system with barcode weighing scale.",
        "Local Commercial Landlord, Equipment Vendors, POS Vendor",
        "Completed physical store, operational hydration wall, functional POS weighing terminal",
        "₹2,50,000 (Store advance deposit, racking, misting, chiller, POS hardware)",
        "Store fit-out completed within 10 days"
    ),
    (
        "Phase 1: Days 21 - 30", "Soft Launch & Grand Opening",
        "1. Staff hiring & training (1 supervisor + 2 sales associates).\n2. First full-scale mandi procurement run.\n3. Inaugural door-to-door pamphlet distribution in 1,500 residential flats.\n4. Grand opening: Offer 'Free 1kg Potato + 250g Onion' on all purchases > ₹199.\n5. Register 100% customer phone numbers at checkout into WhatsApp CRM.",
        "Neighbourhood RWA Societies, Local Residents",
        "Fully operational retail store, 300+ daily walk-ins, database of 500+ customer phone numbers",
        "₹60,000 (Initial opening stock, launch marketing, flyers, staff uniforms)",
        "Daily retail footfall > 250 customers; revenue > ₹25,000/day"
    ),
    (
        "Phase 2: Days 31 - 60", "Offline Stabilization & WhatsApp Rollout",
        "1. Calibrate shrinkage to < 4.5% using 4-tier wastage protocol.\n2. Rollout 'Happy Hour' 30% discount after 7:30 PM.\n3. Launch WhatsApp Business Catalog with daily fresh photos.\n4. Onboard first dedicated delivery executive on EV scooter.\n5. Launch 2 daily delivery slots (Morning 7:30-9:30 AM & Evening 5:30-7:30 PM) for surrounding 3km societies.",
        "WhatsApp Business API, Gated Society Management Committees",
        "Published WhatsApp digital catalog, branded eco-packaging sacks, 25-35 daily online orders",
        "₹40,000 (EV scooter lease advance, Kraft packaging inventory, WhatsApp API setup)",
        "Daily combined sales > ₹35,000; Online delivery accounts for 20% of revenue"
    ),
    (
        "Phase 3: Days 61 - 90", "Full Omnichannel Scale & Web Expansion",
        "1. Launch PWA Hyperlocal Web Store with live inventory sync.\n2. Rollout Weekly Kitchen Box Subscriptions (Standard ₹499/week, Family ₹799/week).\n3. Expand delivery fleet to 2 EV scooters covering 4.5 km radius.\n4. Initiate B2B institutional supply to 5 local premium cafes & restaurants.\n5. Achieve full financial break-even and sustainable net profitability.",
        "5 Local Cafes, Apartment RWAs, Tech Platform Developer",
        "Automated e-commerce storefront, 50+ weekly subscribers, 5 recurring B2B restaurant accounts",
        "₹50,000 (PWA development, digital marketing, B2B crate float)",
        "Daily revenue > ₹45,000 (₹13.5 Lakhs/month); Net Profit > ₹2.4 Lakhs/month"
    )
]

for r_idx, row_item in enumerate(roadmap_data, 5):
    ws6.row_dimensions[r_idx].height = 65
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_item, 1):
        c = ws6.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular if c_idx != 1 else font_bold
        c.fill = fill_zebra if is_even else fill_white
        c.border = border_cell
        c.alignment = align_center if c_idx in [1, 6, 7] else align_left

col_widths6 = {"A": 20, "B": 28, "C": 48, "D": 32, "E": 36, "F": 22, "G": 28}
for col_letter, width in col_widths6.items():
    ws6.column_dimensions[col_letter].width = width

print("Sheet 6 (Phased_Roadmap_30_60_90) generated.")

wb.save(EXCEL_PATH)
print(f"Master Omnichannel Business Model successfully saved to {EXCEL_PATH}!")
