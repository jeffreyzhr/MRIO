import pymrio
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tqdm import tqdm

'''
Data setup --> imported from Z_analysis
Note: this will take a LONG TIME to run. Something like 5 mins per year

TO RUN: navigate to file in terminal and run "python3 main.py"
'''

# Change this list to add or remove years to be included in the spreadsheet
years_list = [2019, 2018, 2017, 2016, 2015, 2014, 2013, 2012, 2011, 2010, 2009, 2008, 2007, 2006, 2005, 2004, 2003, 2002, 2001, 2000]
all_sectors = ['Paddy rice', 'Wheat', 'Cereal grains nec', 'Vegetables, fruit, nuts',
       'Oil seeds', 'Sugar cane, sugar beet', 'Plant-based fibers',
       'Crops nec', 'Cattle', 'Pigs', 'Poultry', 'Meat animals nec',
       'Animal products nec', 'Raw milk', 'Wool, silk-worm cocoons',
       'Manure (conventional treatment)', 'Manure (biogas treatment)',
       'Products of forestry, logging and related services (02)',
       'Fish and other fishing products; services incidental of fishing (05)',
       'Anthracite', 'Coking Coal', 'Other Bituminous Coal',
       'Sub-Bituminous Coal', 'Patent Fuel', 'Lignite/Brown Coal',
       'BKB/Peat Briquettes', 'Peat',
       'Crude petroleum and services related to crude oil extraction, excluding surveying',
       'Natural gas and services related to natural gas extraction, excluding surveying',
       'Natural Gas Liquids', 'Other Hydrocarbons',
       'Uranium and thorium ores (12)', 'Iron ores',
       'Copper ores and concentrates', 'Nickel ores and concentrates',
       'Aluminium ores and concentrates',
       'Precious metal ores and concentrates',
       'Lead, zinc and tin ores and concentrates',
       'Other non-ferrous metal ores and concentrates', 'Stone',
       'Sand and clay',
       'Chemical and fertilizer minerals, salt and other mining and quarrying products n.e.c.',
       'Products of meat cattle', 'Products of meat pigs',
       'Products of meat poultry', 'Meat products nec',
       'products of Vegetable oils and fats', 'Dairy products',
       'Processed rice', 'Sugar', 'Food products nec', 'Beverages',
       'Fish products', 'Tobacco products (16)', 'Textiles (17)',
       'Wearing apparel; furs (18)', 'Leather and leather products (19)',
       'Wood and products of wood and cork (except furniture); articles of straw and plaiting materials (20)',
       'Wood material for treatment, Re-processing of secondary wood material into new wood material',
       'Pulp',
       'Secondary paper for treatment, Re-processing of secondary paper into new pulp',
       'Paper and paper products', 'Printed matter and recorded media (22)',
       'Coke Oven Coke', 'Gas Coke', 'Coal Tar', 'Motor Gasoline',
       'Aviation Gasoline', 'Gasoline Type Jet Fuel', 'Kerosene Type Jet Fuel',
       'Kerosene', 'Gas/Diesel Oil', 'Heavy Fuel Oil', 'Refinery Gas',
       'Liquefied Petroleum Gases (LPG)', 'Refinery Feedstocks', 'Ethane',
       'Naphtha', 'White Spirit & SBP', 'Lubricants', 'Bitumen',
       'Paraffin Waxes', 'Petroleum Coke', 'Non-specified Petroleum Products',
       'Nuclear fuel', 'Plastics, basic',
       'Secondary plastic for treatment, Re-processing of secondary plastic into new plastic',
       'N-fertiliser', 'P- and other fertiliser', 'Chemicals nec', 'Charcoal',
       'Additives/Blending Components', 'Biogasoline', 'Biodiesels',
       'Other Liquid Biofuels', 'Rubber and plastic products (25)',
       'Glass and glass products',
       'Secondary glass for treatment, Re-processing of secondary glass into new glass',
       'Ceramic goods',
       'Bricks, tiles and construction products, in baked clay',
       'Cement, lime and plaster',
       'Ash for treatment, Re-processing of ash into clinker',
       'Other non-metallic mineral products',
       'Basic iron and steel and of ferro-alloys and first products thereof',
       'Secondary steel for treatment, Re-processing of secondary steel into new steel',
       'Precious metals',
       'Secondary preciuos metals for treatment, Re-processing of secondary preciuos metals into new preciuos metals',
       'Aluminium and aluminium products',
       'Secondary aluminium for treatment, Re-processing of secondary aluminium into new aluminium',
       'Lead, zinc and tin and products thereof',
       'Secondary lead for treatment, Re-processing of secondary lead into new lead',
       'Copper products',
       'Secondary copper for treatment, Re-processing of secondary copper into new copper',
       'Other non-ferrous metal products',
       'Secondary other non-ferrous metals for treatment, Re-processing of secondary other non-ferrous metals into new other non-ferrous metals',
       'Foundry work services',
       'Fabricated metal products, except machinery and equipment (28)',
       'Machinery and equipment n.e.c. (29)',
       'Office machinery and computers (30)',
       'Electrical machinery and apparatus n.e.c. (31)',
       'Radio, television and communication equipment and apparatus (32)',
       'Medical, precision and optical instruments, watches and clocks (33)',
       'Motor vehicles, trailers and semi-trailers (34)',
       'Other transport equipment (35)',
       'Furniture; other manufactured goods n.e.c. (36)',
       'Secondary raw materials',
       'Bottles for treatment, Recycling of bottles by direct reuse',
       'Electricity by coal', 'Electricity by gas', 'Electricity by nuclear',
       'Electricity by hydro', 'Electricity by wind',
       'Electricity by petroleum and other oil derivatives',
       'Electricity by biomass and waste', 'Electricity by solar photovoltaic',
       'Electricity by solar thermal', 'Electricity by tide, wave, ocean',
       'Electricity by Geothermal', 'Electricity nec',
       'Transmission services of electricity',
       'Distribution and trade services of electricity', 'Coke oven gas',
       'Blast Furnace Gas', 'Oxygen Steel Furnace Gas', 'Gas Works Gas',
       'Biogas', 'Distribution services of gaseous fuels through mains',
       'Steam and hot water supply services',
       'Collected and purified water, distribution services of water (41)',
       'Construction work (45)',
       'Secondary construction material for treatment, Re-processing of secondary construction material into aggregates',
       'Sale, maintenance, repair of motor vehicles, motor vehicles parts, motorcycles, motor cycles parts and accessoiries',
       'Retail trade services of motor fuel',
       'Wholesale trade and commission trade services, except of motor vehicles and motorcycles (51)',
       'Retail  trade services, except of motor vehicles and motorcycles; repair services of personal and household goods (52)',
       'Hotel and restaurant services (55)', 'Railway transportation services',
       'Other land transportation services',
       'Transportation services via pipelines',
       'Sea and coastal water transportation services',
       'Inland water transportation services', 'Air transport services (62)',
       'Supporting and auxiliary transport services; travel agency services (63)',
       'Post and telecommunication services (64)',
       'Financial intermediation services, except insurance and pension funding services (65)',
       'Insurance and pension funding services, except compulsory social security services (66)',
       'Services auxiliary to financial intermediation (67)',
       'Real estate services (70)',
       'Renting services of machinery and equipment without operator and of personal and household goods (71)',
       'Computer and related services (72)',
       'Research and development services (73)',
       'Other business services (74)',
       'Public administration and defence services; compulsory social security services (75)',
       'Education services (80)', 'Health and social work services (85)',
       'Food waste for treatment: incineration',
       'Paper waste for treatment: incineration',
       'Plastic waste for treatment: incineration',
       'Intert/metal waste for treatment: incineration',
       'Textiles waste for treatment: incineration',
       'Wood waste for treatment: incineration',
       'Oil/hazardous waste for treatment: incineration',
       'Food waste for treatment: biogasification and land application',
       'Paper waste for treatment: biogasification and land application',
       'Sewage sludge for treatment: biogasification and land application',
       'Food waste for treatment: composting and land application',
       'Paper and wood waste for treatment: composting and land application',
       'Food waste for treatment: waste water treatment',
       'Other waste for treatment: waste water treatment',
       'Food waste for treatment: landfill', 'Paper for treatment: landfill',
       'Plastic waste for treatment: landfill',
       'Inert/metal/hazardous waste for treatment: landfill',
       'Textiles waste for treatment: landfill',
       'Wood waste for treatment: landfill',
       'Membership organisation services n.e.c. (91)',
       'Recreational, cultural and sporting services (92)',
       'Other services (93)', 'Private households with employed persons (95)',
       'Extra-territorial organizations and bodies']


df = pd.DataFrame(data=all_sectors)
df_total = {}


def save(database, lst, year):
    database[year] = lst
    return database

count = 1
for i in tqdm(range(len(years_list)), desc="Years completed"):
    year = years_list[i]
    
    path_shorthand = f'/Users/jeffreyzhou/Desktop/MRIO/Data/IOT_{year}_pxp.zip'
    exio3 = pymrio.parse_exiobase3(path=path_shorthand)
    all_regions = exio3.get_regions()

    #include all missing calculations with calc_all()
    exio3.calc_all()
    
    # Consumption impacts
    imports = exio3.impacts.D_imp.US.loc['GHG emissions (GWP100) | Problem oriented approach: baseline (CML, 2001) | GWP100 (IPCC, 2007)']
    consumption = exio3.impacts.D_cba.US.loc['GHG emissions (GWP100) | Problem oriented approach: baseline (CML, 2001) | GWP100 (IPCC, 2007)']
    
    sum = imports.sum() + consumption.sum()
    df_total[year] = sum
    
    imports = imports.to_dict()
    consumption = consumption.to_dict()
    
    
    impacts_percent= {}
    count = 0
    for sector in all_sectors:
        d = consumption[sector]
        n = imports[sector]
        if d == 0:
            count += 1
            if n != 0:
                print("ERROR")
            impacts_percent[sector] = 'NA'
            continue
        percent = n / (d + n)
        impacts_percent[sector] = percent
    
    x, y = zip(*impacts_percent.items())
    df = save(df, y, year)
    count += 1
    
df.to_excel('consumption_impacts.xlsx')
total = pd.DataFrame(data=df_total, index=['total comsumption GHG'])

total = (df_total.T)

total.to_excel('total comsumption GHG')

print("Calculations complete.")
print("Years successfully processed:", count)
    


