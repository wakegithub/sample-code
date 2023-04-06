"""
SQLite3 US Census Data Sample
waiky.jung@gmail.com
This program imports US Census data (available at https://data.census.gov/all?q=zip+code),
imports it into a SQLite3 db file, query it for population & median income data by zip code,
averages it out by state, and output into a new csv file.
"""
import sqlite3
import csv
import re

db = sqlite3.connect('uscensus.db')
db.execute('''DROP TABLE IF EXISTS state_zip_codes''')
db.execute('''DROP TABLE IF EXISTS age_sex''')
db.execute('''DROP TABLE IF EXISTS income''')

def CleanName(table_name):
    temp = re.sub('[^a-zA-Z0-9 ]', ' ', table_name)
    temp = re.sub(' +', ' ', temp)
    temp = re.sub(' ', '_', temp)
    temp = temp.strip()
    return temp

def FixTableNames(fields):
    new_fields = []
    for field in fields:
        field = CleanName(field)
        if field not in new_fields:
            new_fields.append(field)
        else:
            x = '_'
            while field + '_' + str(x) in new_fields:
                x += str('_')
            new_fields.append(field + str(x))

    return new_fields

def MakeTable(census_file, table_name):
    with open(census_file, 'r') as f:
        file = csv.reader(f)
        if 'ACSST5Y' in census_file:
            file.__next__()
        fields = next(file)
        fields = FixTableNames(fields)
        fields_combined = ', '.join(fields)
        if 'ACSST5Y' in census_file:
            fields_combined += 'blank'

        create_table = 'CREATE TABLE IF NOT EXISTS ' + table_name + ' (' + fields_combined + ')'
        db.execute(create_table)

        columns = next(file)
        query = 'insert into ' + table_name + ' ({0}) values ({1})'
        query = query.format(fields_combined, ','.join('?' * len(columns)))
        cursor = db.cursor()
        for data in file:
            data = [0 if x == 'null' else x for x in data]
            cursor.execute(query, data)
        db.commit()

MakeTable('state_zip_codes.csv', 'state_zip_codes')
MakeTable('ACSST5Y2021.S0101-Data.csv', 'age_sex')
MakeTable('ACSST5Y2021.S1901-Data.csv', 'income')

cursor = db.cursor()
cursor.execute('select substr(geographic_area_name, -5) as \'zip_code\', Estimate_Total_Total_population from age_sex')
population = cursor.fetchmany(10)
cursor.execute('select substr(geographic_area_name, -5) as \'zip_code\', Estimate_Households_Mean_income_dollars_ from income')
income = cursor.fetchmany(10)

state_query = """
with previous_query as (
select substr(age_sex.Geographic_Area_Name, -5) as 'zip_code', 
	age_sex.Estimate_Total_Total_population as 'total_population', 
	income.Estimate_Households_Median_income_dollars_ as 'median_income'
from age_sex
join income 
on age_sex.Geographic_Area_Name = income.Geographic_Area_Name)
select state_zip_codes.state, round(avg(previous_query.total_population),1) as 'avg_pop_per_zip_code', round(avg(previous_query.median_income),1) as 'avg_median_income_per_zip_code'
from state_zip_codes
join previous_query
on state_zip_codes.zip_code = previous_query.zip_code
group by state_zip_codes.state
"""
cursor.execute(state_query)
state_avg_pop_and_income_per_zip_code = cursor.fetchall()

with open('output.csv', 'w', newline='') as output:
    fields = ['State', 'Avg. Population per Zip Code', 'Avg. Median Income per Zip Code']
    csv_writer = csv.writer(output)
    csv_writer.writerow(fields)
    csv_writer.writerows(state_avg_pop_and_income_per_zip_code)

print('Done!')