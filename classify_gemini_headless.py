#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import time

from openpyxl import load_workbook
from google import genai
from google.genai import types

INPUT_RATE = 0.30 / 1_000_000   # $ per input token
OUTPUT_RATE = 2.5 / 1_000_000   # $ per output token
AUTOSAVE_EVERY = 100  # save workbook every 100 processed rows

SYSTEM_PROMPT = """You are an expert occupation and business classifier. You will receive in a single request: 1) APPLICANT_NAME and 2) CAM_COMMENT -> Full CAM narrative for the loan file (may contain multiple persons). First, read and understand the CAM_COMMENT completely, identify all persons mentioned, and for each person infer their occupation or business activity and the company or business they work in, work for, run, or operate, along with that company’s business activity. Then focus ONLY on APPLICANT_NAME: match APPLICANT_NAME against the persons you identified using case-insensitive and partial matching (e.g. "Boopathi Raja" can match "Boopathi", "Raja", or "Boopathi Raja"); if multiple people have similar names, prefer the one that best matches the full APPLICANT_NAME string; if you cannot find any clear match, pick the person whose occupation or business is most likely to correspond to that APPLICANT_NAME given the context. Now classify that applicant into EXACTLY ONE label for each of the following, chosen ONLY from the provided catalog (exact text): "industry", "business_category", and "business_profile". VERY IMPORTANT: Industry / Business Category / Business Profile must always be assigned from the POV of the company the applicant is associated with, not their individual job title; that is, these three fields must describe the business of the company they are running, working in, working for, or operating — for example, if the person is an employee in a company producing films, the correct classification should reflect a film/TV/advertising/marketing type company (e.g. Advertising / Marketing / Film/TV Industry at the industry level, and the appropriate film/TV production business category/profile), even if their personal role is just cameraman or editor. All category values in the catalog are separated by commas; select the closest and most specific possible match. If no reasonable match exists in the catalog, you MUST use: { "industry": "Unknown", "business_category": "Unknown", "business_profile": "Unknown", "summary": "No clear occupation or business information available for this applicant." }. Otherwise, return a SHORT structured summary (1–2 sentences) describing ONLY APPLICANT_NAME’s occupation/business in terms of the company’s business activity. OUTPUT FORMAT : { "industry": "...", "business_category": "...", "business_profile": "...", "summary": "Short structured summary about APPLICANT_NAME only, based on the company’s line of business." } 
You MUST respond ONLY with a single valid JSON object exactly in this OUTPUT FORMAT, with double quotes and no additional commentary, markdown, or code fences.
CATALOG (Industry>Business Category>Business Profile): 
Accounting & Auditing > Service Provider >Accounting/ Auditing Services,
Accounting & Auditing >Service Provider >Tax Services,
Advertising / Marketing / Film / TV Industry > Service Provider > Advertising / Marketing Agency,
Advertising / Marketing / Film / TV Industry > Service Provider > Freelancer Marketer,
Advertising / Marketing / Film / TV Industry > Service Provider > Model / Actor / Influencer,
Advertising / Marketing / Film / TV Industry > Service Provider > Production House,
Advertising / Marketing / Film / TV Industry > Service Provider >Media/Entertainment,
Agriculture equipment's > Trader/Retailer > Agriculture equipment's Trader / dealer,
Automobiles / Auto-ancillaries > Trader/Retailer > Accessories,
Automobiles / Auto-ancillaries > Manufacturer > Automobiles / Auto-ancillaries / Accessories Company,
Automobiles / Auto-ancillaries > Trader/Retailer > Batteries,
Automobiles / Auto-ancillaries > Service Provider > Denting and Painting Works,
Automobiles / Auto-ancillaries > Trader/Retailer > Distributor / Whole Seller / Dealership,
Automobiles / Auto-ancillaries > Service Provider > Mechanic (With Setup),
Automobiles / Auto-ancillaries > Service Provider > Mechanic (Without Setup),
Automobiles / Auto-ancillaries > Trader/Retailer > Old Car Sale & Purchase,
Automobiles / Auto-ancillaries > Service Provider > Puncher (With Setup),
Automobiles / Auto-ancillaries > Service Provider > Puncher (Without Setup),
Automobiles / Auto-ancillaries > Service Provider > Service Centre,
Automobiles / Auto-ancillaries > Trader/Retailer > Spare Part Shop,
Automobiles / Auto-ancillaries > Service Provider > Tyre business and service centre,
Automobiles / Auto-ancillaries > Trader/Retailer > Tyre business and service centre,
Aviation > Service Provider >Airlines Company,
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Beauty Parlour (Business Setup),
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Beauty Parlour (Service on demand) / Operations without Set-up,
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Coaching Centre,
Beauty Parlour / Hair Saloon / Tattoo Parlour > Trader/Retailer > Cosmetic Products / Accessories & Other Supplies,
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Hair Saloon (with Business Setup),
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Hair Saloon / Tattoo (Street Hawker),
Beauty Parlour / Hair Saloon / Tattoo Parlour > Trader/Retailer > Saloon / Beauty parlour material,
Beauty Parlour / Hair Saloon / Tattoo Parlour > Service Provider > Tattoo Parlour,
BPO > Service Provider >Outsourcing Services,
Brooms Traders (All House Cleaning item hand made) > Manufacturer > Broom / Rope Manufacturers,
Brooms Traders (All House Cleaning item hand made) > Trader/Retailer > Broom / Rope Traders,
Cable TV Operator and Video Parlor > Service Provider > Cable TV & Internet,
Cable TV Operator and Video Parlor > Service Provider > Video / Game Parlour,
Cartons > Manufacturer >Carton Maker/Manufacturer,
Ceramics> Service Provider > Job work in Ceramics, Sanitary ware & Crockery,
Charcoal Manufacturing > Service Provider >Charcoal Manufacturing Company,
Charitable Organization > Service Provider >Charitable Organization,
Chemical & Fertilizer > Manufacturer > Chemical production,
Chemical & Fertilizer > Trader/Retailer > Fertilizers Shop,
Chemical & Fertilizer > Trader/Retailer > Pesticides Shop,
Chemical & Fertilizer > Service Provider >Laboratory,
Clothing & Fashion > Trader/Retailer > Accessories,
Clothing & Fashion > Service Provider > Boutique / Designer,
Clothing & Fashion > Trader/Retailer > Garments (Hawker),
Clothing & Fashion > Manufacturer > Garments / Textiles,
Clothing & Fashion > Trader/Retailer > Garments / Textiles shop,
Clothing & Fashion > Service Provider > Job Work,
Clothing & Fashion > Service Provider > Tailoring,
Construction Industry > Manufacturer > Brick Maker,
Construction Industry > Trader/Retailer > Building Material supplier with Setup,
Construction Industry > Trader/Retailer > Building Material supplier without Setup,
Construction Industry > Trader/Retailer > Ceramics, Sanitary ware & Crockery,
Construction Industry > Manufacturer > Ceramics, Sanitary ware & Crockery manufacturing setup,
Construction Industry > Service Provider > Civil Contractor,
Construction Industry > Service Provider > Construction Equipment Rental,
Construction Industry > Trader/Retailer > Dealer in Construction Equipment,
Construction Industry > Service Provider > Electrician / electrical contractor,
Construction Industry > Service Provider > Electrician / Plumber With Business Setup,
Construction Industry > Manufacturer > Hardware / Sanitary / Paint / Varnish Items,
Construction Industry > Trader/Retailer > Hardware / Sanitary / Paint / Varnish Items,
Construction Industry > Service Provider > Labour Contractor,
Construction Industry > Trader/Retailer > Marble / Granite,
Construction Industry > Service Provider > Mason / Bricklayer,
Construction Industry > Service Provider > Painter / Painting Contractor,
Construction Industry > Service Provider > Plumber / Plumber Contractor,
Corporate Services > Service Provider >Executive/Manager,
Corporate Services > Service Provider >HR / Recruitment Services ,
Corporate Services > Service Provider >Sales and Field Marketing Services ,
Courier > Service Provider > Courier Franchisee,
Courier > Service Provider >Delivery Agent,
Customer Essential Services > Service Provider > Common Service Centre,
Cycle Repairing & Spare Parts > Trader/Retailer > Sale & Repairing Shop,
Defence> Service Provider > Armed Forces / Fire / Rescue services,
Diamond > Service Provider > Broker,
Diamond > Trader/Retailer > Diamond Dealer,
Diamond > Service Provider > Diamond Polish Work,
Diamond > Service Provider > Job Worker,
Education > Service Provider > Art / Music / Dance / Karate Classes,
Education > Service Provider > Coaching Centre,
Education > Service Provider > Consultancy services for students aspiring education overseas,
Education > Service Provider > Counsellor,
Education > Service Provider > Library,
Education > Service Provider > Private Tuitions / tutors,
Education > Service Provider > Training Institute,
Education > Service Provider > Typing / Shorthand / basic Computers,
Education > Service Provider >School/College/Educational Institution,
Electronic / Home Appliance (TV, Fridge, Mixer Grinder etc) > Trader/Retailer > Distributor,
Electronic / Home Appliance (TV, Fridge, Mixer Grinder etc) > Manufacturer > Electronic / Home Appliance,
Electronic / Home Appliance (TV, Fridge, Mixer Grinder etc) > Trader/Retailer > Electronics shop (sale),
Electronic / Home Appliance (TV, Fridge, Mixer Grinder etc) > Service Provider > Service & Repairs,
Electronics, IT & Telecom > Manufacturer > Computer / Mobile,
Electronics, IT & Telecom > Trader/Retailer > Computer / Mobile Distributor,
Electronics, IT & Telecom > Trader/Retailer > Computer / Mobile Shop,
Electronics, IT & Telecom > Service Provider >IT Services,
Elevators & Lift > Service Provider > Installation and maintenance services,
Elevators & Lift > Trader/Retailer >Elevators & Lift Manufacturing Company,
Entertainment and Events Service > Service Provider > Artist,
Event Management > Service Provider > Decorators,
Event Management > Service Provider > DJ / Sound Services,
Event Management > Service Provider > Event Organiser,
Event Management > Service Provider > Marriage Hall,
Event Management > Service Provider > Rentals (chair, vessels, lights, speakers, Tent etc),
Farming & Animal Husbandry > Trader/Retailer > Aquarium,
Farming & Animal Husbandry > Manufacturer > Bamboo business / related products,
Farming & Animal Husbandry > Manufacturer > Farming of Crops / Vegetables,
Farming & Animal Husbandry > Manufacturer > Farming of Crops / Vegetables,
Farming & Animal Husbandry > Manufacturer > Fish farming,
Farming & Animal Husbandry > Manufacturer > Goat farming,
Farming & Animal Husbandry > Trader/Retailer > Livestock Trader,
Farming & Animal Husbandry > Manufacturer > Mushroom seeds and trading,
Farming & Animal Husbandry > Manufacturer > Pet farming,
Farming & Animal Husbandry > Trader/Retailer > Pet Shop,
Farming & Animal Husbandry > Manufacturer > Poultry Farm,
Farming & Animal Husbandry > Manufacturer > Sheep farming,
Farming & Animal Husbandry > Service Provider > Wholesale and Distribution,
Financial Services - Investment / Consultancy > Service Provider > Currency Dealer,
Financial Services - Investment / Consultancy > Service Provider > Post Office Agent,
Financial Services - Investment / Consultancy > Service Provider >Bank/Insurance/Investment Company,
Financial Services - Investment / Consultancy > Service Provider >Recovery Services / Agent,
Flower shop > Trader/Retailer > Car / Marriage / Party Decorator,
Flower shop > Trader/Retailer > Flower Shop with Permanent setup,
Flower shop > Trader/Retailer > Street hawker,
FMCG > Manufacturer > Bidi Works / Tobacco,
FMCG > Trader/Retailer > Oil Mill,
FMCG > Trader/Retailer > Res-cum-business with small machines for individual households - Flour Mill,
FMCG >Trader/Retailer>FMCG Company ,
Food and Beverage Sector > Service Provider > Bakery Shop,
Food and Beverage Sector > Service Provider > Catering,
Food and Beverage Sector > Trader/Retailer > Chicken / Mutton Shop,
Food and Beverage Sector > Trader/Retailer > Dealing in spices,
Food and Beverage Sector > Trader/Retailer > Distributor / Wholesaler,
Food and Beverage Sector > Service Provider > Fast food centre,
Food and Beverage Sector > Trader/Retailer > Fish Shop,
Food and Beverage Sector > Trader/Retailer > Flour Mill with business setup,
Food and Beverage Sector > Service Provider > Food processing,
Food and Beverage Sector > Service Provider > Home Made Food Product / Papad Making,
Food and Beverage Sector > Service Provider > Ice Cream Parlour & Juice,
Food and Beverage Sector > Trader/Retailer > Kirana shop / general store,
Food and Beverage Sector > Manufacturer > Manufacturer of spices,
Food and Beverage Sector > Trader/Retailer > PAN parlour, cold drink and General,
Food and Beverage Sector > Service Provider > Restaurant / Dhaba,
Food and Beverage Sector > Service Provider > Street Food Hawker,
Food and Beverage Sector > Service Provider > Sweet / Snack Shop,
Food and Beverage Sector > Service Provider > Tea stall,
Food and Beverage Sector > Service Provider > Tiffin Centre & Mess,
Food and Beverage Sector > Service Provider >Food/Beverage  Company ,
Footwear > Trader/Retailer > Footwear (Hawker),
Footwear > Trader/Retailer > Footwear shop,
Footwear > Service Provider > Job work,
Footwear > Service Provider > Mochi / Footwear Repair / Cobler,
Fruits / Vegetables Vendor > Trader/Retailer > Fruit / Vegetable supplier / wholesaler,
Fruits / Vegetables Vendor > Trader/Retailer > Fruit / Vegetable Vendor (Hawker),
Glass works > Service Provider > Glass & Glass Products,
Glass works > Trader/Retailer > Glass & Glass Products,
Government / Public Services > Service Provider > Charge / Telecom Mechanic,
Government / Public Services > Service Provider >Government / Public Services ,
Government / Public Services > Service Provider >Railways Department,
Gym / Wellness Centre > Service Provider > Gym Franchisee,
Gym / Wellness Centre > Service Provider > Gym Trainer,
Gym / Wellness Centre > Service Provider > Local Gym,
Gym / Wellness Centre > Service Provider > Yoga / Naturopathy,
Handicraft / Handloom > Service Provider > Handicraft,
Handicraft / Handloom > Service Provider > Handloom,
Handicraft / Handloom > Service Provider > Sculptor,
Handicraft / Handloom > Trader/Retailer > Stall at pilgrim centres,
Health Care > Service Provider > Ayurvedic / Homeopathic / Unani Doctor,
Health Care > Trader/Retailer > Ayurvedic / Homeopathic Medical & General Shop,
Health Care > Service Provider > Diagnostic Services,
Health care > Service Provider > Doctor,
Health Care > Service Provider > First Aid centre,
Health Care > Manufacturer > Medical / Herbal Product,
Health Care > Trader/Retailer > Medical / Herbal Product,
Health Care > Manufacturer > Medical Equipment,
Health Care > Trader/Retailer > Medical Equipment,
Health Care > Service Provider > Scientist,
Health Care > Service Provider >Pharmaceuticals ,
Hospitality > Service Provider >Chef and Hospitality Services,
Hotel / Guest House > Service Provider > Guest House / PG,
Hotel / Guest House > Service Provider > Hotel with boarding / lodging facility,
House Keeping > Trader/Retailer > House Cleaning & Wash Materials,
House Keeping > Service Provider > House Keeping Services provider,
Interior Designer / Architect > Service Provider > Architect,
Interior Designer / Architect > Service Provider > Interior Designer,
Jewellery - Imitation / Gold / Silver / Stones > Trader/Retailer > Gold and Silver shop,
Jewellery - Imitation / Gold / Silver / Stones > Manufacturer > Imitation Jewellery,
Jewellery - Imitation / Gold / Silver / Stones > Trader/Retailer > Imitation Jewellery,
Jewellery - Imitation / Gold / Silver / Stones > Service Provider > Job Worker,
Kites Business > Manufacturer > Kite Maker / Manufacturer,
Kites Business > Trader/Retailer > Kite Trader,
Landlord (Rental Income) > Service Provider > Commercial property,
Landlord (Rental Income) > Service Provider > Residential property,
Laundry & Dry Cleaners > Service Provider > Dhobi,
Laundry & Dry Cleaners > Service Provider > Laundry / Dry Cleaner,
Legal Related Services > Service Provider > Deed Writer / typist / Notary work,
Luggage / Bag > Manufacturer > Bag maker / Bag factory,
Luggage / Bag > Service Provider > Bag repair,
Luggage / Bag > Trader/Retailer > Bag Seller Hawker,
Luggage / Bag > Trader/Retailer > Bag Seller with Business setup,
Metal work > Trader/Retailer > Aluminium work & Fabrication,
Metal work > Service Provider > Engg workshop / Welding / Moulding,
Metal work > Service Provider > Job worker,
Metal work > Trader/Retailer > Metal Products,
Metal work > Service Provider > Metal work,
Milk Business / Dairy > Service Provider > Agency / Collection / Distributor,
Milk Business / Dairy > Service Provider > Cattle Dairy Business,
Milk Business / Dairy > Service Provider > Milk Dairy shops,
Mining > Manufacturer > Mining and Tunneling,
Mining > Service Provider > Upstream Drilling / Rig Operations,
Musical instrument > Trader/Retailer > Musical instrument dealer,
Musical instrument > Manufacturer > Musical instrument manufacturer,
News Agency > Service Provider > News Paper Agency,
News Paper Vendor > Trader/Retailer > News Paper Stall,
Nursery / Plant Business > Manufacturer > Owning Farm Lands,
Nursery / Plant Business > Trader/Retailer > Sales point for Nursery / Plant business,
Packaging and Disposable Items> Trader/Retailer > Packaging and Disposable Items,
Packaging and Disposable Items > Manufacturer > Packaging and Disposable Items,
Packaging and Disposable Items > Trader/Retailer > Packaging and Disposable Items,
Painting > Service Provider > Painter / Painting Contractor,
Paper Manufacturing > Service Provider >Paper Manufacturing Company,
Pest Control > Service Provider > Pest Control Services with Business Setup,
Pest Control > Service Provider > Pest Control Services without Business Setup,
Petrol Pump & Gas Agency > Trader/Retailer > CNG Station,
Petrol Pump & Gas Agency > Trader/Retailer > Gas Agency,
Petrol Pump & Gas Agency > Trader/Retailer > LPG Dealers,
Petrol Pump & Gas Agency > Trader/Retailer > Petrol Pump,
Photography / Digital Studio > Service Provider > Music Recording Studio,
Photography / Digital Studio > Service Provider > Photographer,
Photography / Digital Studio > Service Provider > Photography shop / studio,
Plastic Items > Manufacturer > Trading in plastic Items,
Plastic Items > Trader/Retailer > Trading in plastic Items,
Potter (Matke wala ) manufacture and sale > Manufacturer > Potter (Matka Maker),
Potter (Matke wala ) manufacture and sale > Trader/Retailer > Potter (Matka Trader),
Power Generation > Service Provider >Power Generation Company,
Power Loom > Service Provider > Job worker,
Power Loom > Manufacturer > Power Loom factory,
Printing Press / Graphics > Service Provider > Graphic Designer,
Printing Press / Graphics > Service Provider > Printing Press,
Printing Press / Graphics > Service Provider > Printing Services,
Professional Services - CA / CS / ICWA / Advocate > Service Provider > Advocate,
Professional Services - CA / CS / ICWA / Advocate > Service Provider > CA,
Professional Services - CA / CS / ICWA / Advocate > Service Provider > CS,
Real Estate > Service Provider >Real Estate Agent,
Religious Services & Products > Service Provider > Astrologer / Vastu Consultant,
Religious Services & Products > Service Provider > Religious / Pandit Services,
Religious Services & Products > Manufacturer > Religious Products (Pooja samagri, etc),
Religious Services & Products > Trader/Retailer > Religious Products (Pooja samagri, etc),
Rental> Service Provider >Leasehold Services,
Retail > Trader/Retailer > General store - Non food (Cutlery store),
Rubber & rubber products > Manufacturer > Manufacturer of Rubber, rubber products,
Rubber & rubber products > Trader/Retailer > Trading in Rubber, rubber products,
Scrap Dealer > Trader/Retailer > Homebased scrap dealer,
Scrap Dealer > Trader/Retailer > Industrial Scrap dealer,
Scrap Dealer > Trader/Retailer > Plastic Recycle Work,
Security Services > Service Provider > Commercial Security,
Security Services > Service Provider > Fire extinguisher sales & servicing,
Security Services > Service Provider > Residential Security,
Social / Welfare Services > Service Provider >Social / Welfare Services ,
Sports > Manufacturer >Sports Equipment Manufacturer,
Stationary / Gift Items > Manufacturer > Gift Items,
Stationary / Gift Items > Trader/Retailer > Stationary / Gift Shop,
Stationary / Gift Items > Manufacturer > Stationary Items,
Tours & Travels > Service Provider > Driver,
Tours & Travels > Service Provider > Provides Vehicle on Rental / Hire (No-Owned Vehicle),
Tours & Travels > Service Provider > Provides Vehicle on Rental / Hire (Owned Vehicle),
Tours & Travels > Service Provider > Tours & Travels,
Toys & Fancy Items > Trader/Retailer > Toys & Fancy Store,
Transport & Logistics > Service Provider > Commission Agent,
Transport & Logistics > Service Provider > Driver,
Transport & Logistics > Service Provider > Driver (Without Vehicle),
Transport & Logistics > Service Provider > Goods carrier,
Transport & Logistics > Service Provider > Logistics / Export Operations,
Transport & Logistics > Service Provider > Rapido / Zomato / Swiggy - Gig Worker,
Transport & Logistics > Service Provider > Rickshaw Driver (With Vehicle Ownership),
Transport & Logistics > Service Provider > Rickshaw Driver (Without Vehicle Ownership),
Transport & Logistics > Service Provider > School Taxi Services,
Utensils > Trader/Retailer > Utensils & Crockery Shop,
Utensils > Manufacturer > Utensils manufacturing setup,
Warehousing> Service Provider >Storage & Warehousing Services,
Water Supplier > Service Provider > Bore wells,
Water Supplier > Trader/Retailer > Motors, generators and pumps, Transformers dealer,
Water Supplier > Manufacturer > RO Plant owner,
Water Supplier > Trader/Retailer > RO water distributor,
Water Supplier > Service Provider > Water Supplier,
Wood / Furniture / Home Décor > Service Provider > Carpenter / Carpenter Contractor,
Wood / Furniture / Home Décor > Trader/Retailer > Curtains / Carpets / mattress / Pillow Maker,
Wood / Furniture / Home Décor > Manufacturer > Furniture / Timber / Wood Products,
Wood / Furniture / Home Décor > Trader/Retailer > Furniture / Timber / Wood Products"""  # copy-paste exactly

def process_excel(args):
    # GEMINI_API_KEY is read from env
    client = genai.Client()

    print(f"Opening workbook: {args.wb}, sheet: {args.sheet}")
    wb = load_workbook(args.wb)
    sheet = wb[args.sheet]

    name_col = args.name_col.upper()
    comment_col = args.comment_col.upper()

    out_path_col = "D"      # combined: industry > business_category > business_profile
    out_summary_col = "E"   # applicant summary

    # Headers
    def ensure_header(col, text):
        cell = f"{col}1"
        if not (sheet[cell].value or "").strip():
            sheet[cell].value = text

    ensure_header(out_path_col, "Industry > Business Category > Business Profile")
    ensure_header(out_summary_col, "Applicant Summary")

    start_row = args.start_row
    limit = args.limit

    total_prompt_tokens = 0
    total_output_tokens = 0
    rows_processed = 0

    print(
        f"Starting processing from row {start_row} to 10000 "
        f"with limit={limit or 'no limit'}"
    )

    current_row = start_row
    while True:
        # Stop on limit, if set
        if limit and rows_processed >= limit:
            break

        name = sheet[f"{name_col}{current_row}"].value
        comment = sheet[f"{comment_col}{current_row}"].value

        # Stop when both empty
        if (name is None or str(name).strip() == "") and \
           (comment is None or str(comment).strip() == ""):
            break

        # Skip rows with no comment
        if comment is None or str(comment).strip() == "":
            current_row += 1
            continue

        applicant_name = str(name or "").strip()
        cam_comment = str(comment).strip()

        user_prompt = (
            "Use the CATALOG from the system instruction.\n\n"
            f"APPLICANT_NAME:\n{applicant_name}\n\n"
            "CAM_COMMENT:\n"
            f"{cam_comment}\n\n"
            "Remember: focus ONLY on APPLICANT_NAME when returning the classification and summary."
        )

        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0.0,
                    response_mime_type="application/json",
                ),
            )
        except Exception as e:
            print(f"Error at row {current_row}: {e}")
            # write error info into the combined path + summary columns
            sheet[f"{out_path_col}{current_row}"].value = "Error"
            sheet[f"{out_summary_col}{current_row}"].value = str(e)

            rows_processed += 1  # so errors count toward limit/autosave
            if rows_processed % AUTOSAVE_EVERY == 0:
                try:
                    wb.save(args.wb)
                    print(
                        f"[AUTOSAVE] Workbook saved after row {current_row} "
                        f"(total rows processed: {rows_processed})"
                    )
                except Exception as se:
                    print(
                        f"[AUTOSAVE ERROR] Could not save workbook at row {current_row}: {se}"
                    )

            current_row += 1
            continue
        # Extract JSON text from response
        content = ""
        try:
            if hasattr(response, "text") and response.text:
                content = response.text
            else:
                content = response.candidates[0].content.parts[0].text
        except Exception:
            content = ""

        industry = ""
        category = ""
        profile = ""
        summary = ""
        try:
            data = json.loads(content)
            industry = data.get("industry", "") or ""
            category = data.get("business_category", data.get("category", "")) or ""
            profile = data.get("business_profile", data.get("profile", "")) or ""
            summary = data.get("summary", "") or ""
        except Exception:
            industry = "Unknown"
            category = "Unknown"
            profile = "Unknown"
            summary = "Parsing error: could not read model output."

        # Combine as: industry > business_category > business_profile
        combined_path = f"{industry} > {category} > {profile}"

        sheet[f"{out_path_col}{current_row}"].value = combined_path
        sheet[f"{out_summary_col}{current_row}"].value = summary

        usage = getattr(response, "usage_metadata", None)
        if usage is not None:
            pt = getattr(usage, "prompt_token_count", 0) or 0
            total = getattr(usage, "total_token_count", 0) or 0
            ot = max(total - pt, 0)
            total_prompt_tokens += pt
            total_output_tokens += ot
            print(
                f"Row {current_row} done -> prompt_tokens={pt}, output_tokens={ot}, "
                f"industry='{industry}', category='{category}'"
            )
        else:
            print(
                f"Row {current_row} done (no usage_metadata) -> "
                f"industry='{industry}', category='{category}'"
            )

        rows_processed += 1

        # 🔹 Autosave every AUTOSAVE_EVERY rows
        if rows_processed % AUTOSAVE_EVERY == 0:
            try:
                wb.save(args.wb)
                print(
                    f"[AUTOSAVE] Workbook saved after row {current_row} "
                    f"(total rows processed: {rows_processed})"
                )
            except Exception as e:
                print(
                    f"[AUTOSAVE ERROR] Could not save workbook at row {current_row}: {e}"
                )

        current_row += 1
        time.sleep(0.05)

    # Final save
    try:
        wb.save(args.wb)
        print("Excel updated and saved.")
        print(f"Rows processed: {rows_processed}")
    except Exception as e:
        print(f"[FINAL SAVE ERROR] Could not save workbook: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Classify applicant occupations from Excel using Gemini (headless, openpyxl)"
    )

    parser.add_argument(
        "--wb",
        default="Data/AHFL Cases.xlsx",
        help="Path to Excel workbook inside repo (default: Data/AHFL Cases.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default="Sheet1",
        help="Sheet name (default: Sheet1)",
    )
    parser.add_argument(
        "--name_col",
        default="B",
        help="Column letter for applicant name (default: B)",
    )
    parser.add_argument(
        "--comment_col",
        default="C",
        help="Column letter for CAM comment (default: C)",
    )
    parser.add_argument(
        "--start_row",
        type=int,
        default=2,
        help="Row number to start from (default: 2)",
    )
    parser.add_argument(
        "--end_row",
        type=int,
        default=10000,
        help="Last row number to process (inclusive, default: 10000; 0 = no upper limit)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Max rows to process (0 = no limit, used together with start_row / end_row)",
    )

    args = parser.parse_args()
    process_excel(args)
