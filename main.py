import requests
import xlsxwriter
import datetime

class Convertor:
    __flag = False

    def setFrom(self, from_con):
        self.from_con = from_con
    def getFrom(self):
        return self.from_con
    def setTo(self, to_con):
        self.to_con = to_con
    def getTo(self):
        return self.to_con
    def getApi(self):
        return self.api
    def setApi(self, api):
        self.api = api

    def test_connection(self):
        try:
            requests.get('https://www.google.com')
        except:
            return False
        else:
            return True

    def get_country(self):
        try:
            country = requests.get('https://free.currconv.com/api/v7/countries?apiKey=&apiKey={api}'.format(api=self.api))
        except:
            print('Something is wrong.Please try again later.')
            self.__flag = False
        else:
            return country.json()

    def __init__(self, api):
        test = self.test_connection()
        self.setApi(api)
        if test == False:
            print('No internet connection.Please try again.')
        if test == True:
            self.__flag = True
            self.country = self.get_country()

    def convert_all(self, from_con):
        self.setFrom(from_con)
        if self.__flag == False:
            print('Something is wrong.Please try again.')
        if self.__flag == True:
            wb = xlsxwriter.Workbook('convert-from--{base}--{time}.xlsx'.format(base=self.from_con, time=datetime.datetime.now().strftime('%Y-%m-%d-%h-%m')))
            sheet1 = wb.add_worksheet()
            sheet1.write(0, 0, '#')
            sheet1.write(0, 1, 'Name')
            sheet1.write(0, 2, 'CurrencyName')
            sheet1.write(0, 3, 'CurrencySymbol')
            sheet1.write(0, 4, 'Value')
            counter = 1
            for i in self.country:
                for j in self.country[i]:
                    val = requests.get('https://free.currconv.com/api/v7/convert?q={from_con}_{to_con}&apiKey={api}'.format(from_con=self.from_con, to_con=self.country[i][j]['currencyId'], api=self.api))
                    try:
                        val = val.json()['results']['{from_con}_{to_con}'.format(from_con=self.from_con, to_con=self.country[i][j]['currencyId'])]['val']
                    except:
                        print('Something is wrong.Please try again.')
                        break
                    else:
                        sheet1.write(counter, 0, counter)
                        sheet1.write(counter, 1, self.country[i][j]['name'])
                        sheet1.write(counter, 2, self.country[i][j]['currencyName'])
                        sheet1.write(counter, 3, self.country[i][j]['currencySymbol'])
                        sheet1.write(counter, 4, val)
                        counter = counter + 1
            wb.close()

    def convert(self, from_con, to_con):
        self.setFrom(from_con)
        self.setTo(to_con)
        if self.__flag == False:
            print('Something is wrong.Please try again.')
        if self.__flag == True:
            val = requests.get('https://free.currconv.com/api/v7/convert?q={from_con}_{to_con}&apiKey={api}'.format(from_con=self.from_con, to_con=self.to_con, api=self.api))
            try:
                val = val.json()['results']['{from_con}_{to_con}'.format(from_con=from_con, to_con=self.to_con)]['val']
            except:
                print('Something is wrong.Please try again.')
            else:
                print('1 unit of {from_con} equal to {val} of {to_con}'.format(from_con=self.from_con, val=str(val), to_con=self.to_con))

if __name__ == '__main__':
    con = Convertor('edf0254bb75fda82e6f3')
    con.convert('USD', 'IRR')
    con.convert_all('USD')