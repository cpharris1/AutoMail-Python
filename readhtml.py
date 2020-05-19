import csv

resultData = [['jubox79@gmail.com', 'Success', '10-3-2020'], ['jubox79@gmail.com', 'Success', '10-3-2020']]

with open('output.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(resultData)