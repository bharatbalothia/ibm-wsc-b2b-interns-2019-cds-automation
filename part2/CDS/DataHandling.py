from cloudant.client import Cloudant
import uuid
import couchdb



class DataHandling:
    def __init__(self):
        client = Cloudant("e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix", "4299f4e82a5b36181a52abd82d9d74bd5bf3f77a350d5db2daf9b30882df8cb8",
                          url="https://e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix:4299f4e82a5b36181a52abd82d9d74bd5bf3f77a350d5db2daf9b30882df8cb8@e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix.cloudantnosqldb.appdomain.cloud", connect='true')
        server = couchdb.Server("http://%s:%s@9.199.145.193:5984/" % ("admin", "admin123"))
        #self.db = client['cdsdata']
        self.db =server['cdsproject']

    def SaveNewTP(self,tpDetailDict,customerName):
        customer = self.db[customerName]
        tpid = str(uuid.uuid4())
        customer['TPlist'][tpid] = tpDetailDict
        self.db.save(customer)

    def SaveNewCustomer(self,customerDetailDict,customerName):
        if customerName not in self.db:
            self.db.save(customerDetailDict)

    def saveAnyData(self,customerData,tpData):
        customername = customerData['Customer Name']
        self.SaveNewCustomer(customerData,customername)
        self.SaveNewTP(tpData,customername)

    def getCustomerList(self):
        Customers = []
        for data in self.db:
            Customers.append(data)
        return Customers

    def getReport(self,customername):
        tradingpartners = self.db[customername]['TPlist']
        tpDetail = [tradingpartners[id] for id in tradingpartners]
        return tpDetail
    def getTPDetails(self,tpname,customername):
        customer = self.db[customername]
        for id in customer['TPlist']:
            if customer['TPlist'][id]['TP name'] == tpname:
                return customer['TPlist'][id]


    def UpdateTpdetails(self,tpid,tpDetail,customername):
        if customername in self.db:
            customer = self.db[customername]
            customer['TPlist'][tpid] = tpDetail
            self.db.save(customer)
