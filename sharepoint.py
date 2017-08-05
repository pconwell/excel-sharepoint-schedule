from lxml import etree
import requests
from requests_ntlm import HttpNtlmAuth

class Site(object):

    def __init__(self, site_url, auth):
        self.site_url = site_url
        self._session = requests.Session()
        self._session.auth = HttpNtlmAuth(auth[0],auth[1])
        self.last_request = None
        self._services_url = {'Lists':'/_vti_bin/lists.asmx',
                              'Views':'/_vti_bin/Views.asmx'}

    def _url(self, service):
        return ''.join([self.site_url, self._services_url[service]])

    def _headers(self, soapaction):
        headers = {"Content-Type":"text/xml; charset=UTF-8",
                   "SOAPAction":"http://schemas.microsoft.com/sharepoint/soap/" + soapaction}
        return headers

    def List(self, listName):
        return _List(self._session, listName, self._url)


class _List(object):

    def __init__(self, session, listName, url):
        self._session = session
        self.listName = listName
        self._url = url

        # List Info
        self.fields = []
        self.regional_settings = {}
        self.server_settings = {}
        self.GetList()

        self._sp_cols = {i['Name']: {'name': i['DisplayName'], 'type': i['Type']} for i in self.fields}
        self._disp_cols = {i['DisplayName']: {'name': i['Name'], 'type': i['Type']} for i in self.fields}

        title_col = self._sp_cols['Title']['name']
        title_type = self._sp_cols['Title']['type']
        self._disp_cols[title_col] = {'name': 'Title', 'type': title_type}
        # This is a shorter lists that removes the problems with duplicate names for "Title"
        standard_source = 'http://schemas.microsoft.com/sharepoint/v3'

        self.last_request = None

    def _url(self, service):
        return ''.join([self.site_url, self._services_url[service]])

    def _headers(self, soapaction):
        headers = {"Content-Type": "text/xml; charset=UTF-8",
                   "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/" + soapaction}
        return headers

    def _convert_to_internal(self, data):
        """From 'Column Title' to 'Column_x0020_Title'"""
        for _dict in data:
            keys = list(_dict.keys())[:]
            for key in keys:
                if key not in self._disp_cols:
                    raise Exception(key + ' not a column in current List.')
                _dict[self._disp_cols[key]['name']] = self._sp_type(key, _dict.pop(key))

    def _convert_to_display(self, data):
        """From 'Column_x0020_Title' to  'Column Title'"""
        for _dict in data:
            keys = list(_dict.keys())[:]
            for key in keys:
                if key not in self._sp_cols:
                    raise Exception(key + ' not a column in current List.')
                _dict[self._sp_cols[key]['name']] = self._python_type(key, _dict.pop(key))

    def _python_type(self, key, value):
        """Returns proper type from the schema"""
        try:
            field_type = self._sp_cols[key]['type']
            if field_type in ['Number', 'Currency']:
                return float(value)
            elif field_type == 'DateTime':
                # Need to round datetime object
                return datetime.strptime(value, '%Y-%m-%d %H:%M:%S').date()
            elif field_type == 'Boolean':
                if value == '1':
                    return 'Yes'
                elif value == '0':
                    return 'No'
                else:
                    return ''
            elif field_type == 'User':
                # Sometimes the User no longer exists or
                # has a diffrent ID number so we just remove the "123;#"
                # from the beginning of their name
                if value in self.users['sp']:
                    return self.users['sp'][value]
                else:
                    return value.split('#')[1]
            else:
                return value
        except AttributeError:
            return value

    def _sp_type(self, key, value):
        """Returns proper type from the schema"""
        try:
            field_type = self._disp_cols[key]['type']
            if field_type in ['Number', 'Currency']:
                return value
            elif field_type == 'DateTime':
                return value.strftime('%Y-%m-%d %H:%M:%S')
            elif field_type == 'Boolean':
                if value == 'Yes':
                    return '1'
                elif value == 'No':
                    return '0'
                else:
                    raise Exception("%s not a valid Boolean Value, only 'Yes' or 'No'" % value)
            elif field_type == 'User':
                return self.users['py'][value]
            else:
                return value
        except AttributeError:
            return value

    def GetListItems(self, fields):

        # Build Request
        soap_request = soap('GetListItems')
        soap_request.add_parameter('listName', self.listName)

        # Add viewFields
        for i, val in enumerate(fields):
            fields[i] = self._disp_cols[val]['name']
        viewfields = fields
        soap_request.add_view_fields(fields)

        soap_request.add_query({'OrderBy': ['ID']})

        # Send Request
        response = self._session.post(url=self._url('Lists'),
                                     headers = self._headers('GetListItems'),
                                     data = str(soap_request),
                                     verify = True)


        # Parse Response
        if response.status_code == 200:
            envelope = etree.fromstring(response.text.encode('utf-8'))
            listitems = envelope[0][0][0][0][0]
            data = []
            for row in listitems:
                # Strip the 'ows_' from the beginning with key[4:]
                data.append({key[4:]: value for (key, value) in row.items() if key[4:] in viewfields})

            self._convert_to_display(data)

            return data
        else:
            return response

    def GetList(self):

        soap_request = soap('GetList')
        soap_request.add_parameter('listName', self.listName)
        self.last_request = str(soap_request)

        # Send Request
        response = self._session.post(url=self._url('Lists'),
                                     headers = self._headers('GetList'),
                                     data = str(soap_request),
                                     verify = True)

        # Parse Response
        if response.status_code == 200:
            envelope = etree.fromstring(response.text.encode('utf-8'))
            _list = envelope[0][0][0][0]
            info = {key: value for (key,value) in _list.items()}
            for row in _list[0].getchildren():
                self.fields.append({key: value for (key,value) in row.items()})

            for setting in _list[1].getchildren():
                self.regional_settings[setting.tag.strip('{http://schemas.microsoft.com/sharepoint/soap/}')] = setting.text

            for setting in _list[2].getchildren():
                self.server_settings[setting.tag.strip('{http://schemas.microsoft.com/sharepoint/soap/}')] = setting.text
            fields = envelope[0][0][0][0][0]

        else:
            raise Exception("ERROR:", response.status_code, response.text)

    def UpdateListItems(self, data, kind):

        soap_request = soap('UpdateListItems')
        soap_request.add_parameter('listName', self.listName)
        self._convert_to_internal(data)
        soap_request.add_actions(data, kind)
        self.last_request = str(soap_request)

        # Send Request
        response = self._session.post(url=self._url('Lists'),
                                      headers = self._headers('UpdateListItems'),
                                      data = str(soap_request),
                                      verify = True
                                      )

class soap(object):
    """A simple class for building SAOP Requests"""
    def __init__(self, command):
        self.envelope = None
        self.command = command
        self.request = None
        self.updates = None
        self.batch = None

        # HEADER GLOBALS
        SOAPENV_NAMESPACE = "http://schemas.xmlsoap.org/soap/envelope/"
        SOAPENV = "{%s}" % SOAPENV_NAMESPACE
        ns0_NAMESPACE = "http://schemas.xmlsoap.org/soap/envelope/"
        ns0 = "{%s}" % ns0_NAMESPACE
        ns1_NAMESPACE = "http://schemas.microsoft.com/sharepoint/soap/"
        ns1 = "{%s}" % ns1_NAMESPACE
        xsi_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance"
        xsi = "{%s}" % xsi_NAMESPACE
        NSMAP = {'SOAP-ENV' : SOAPENV_NAMESPACE,'ns0':ns0_NAMESPACE,'ns1':ns1_NAMESPACE,'xsi':xsi_NAMESPACE}

        # Create Header
        self.envelope = etree.Element(SOAPENV + "Envelope", nsmap=NSMAP)
        header = etree.SubElement(self.envelope, SOAPENV + "Header", nsmap=NSMAP)
        HEADER = etree.SubElement(self.envelope, '{http://schemas.xmlsoap.org/soap/envelope/}Body')

        # Create Command
        self.command = etree.SubElement(HEADER, '{http://schemas.microsoft.com/sharepoint/soap/}' + command)

        self.start_str = b"""<?xml version="1.0" encoding="utf-8"?>"""

    def add_parameter(self, parameter, value=None):
        sub = etree.SubElement(self.command, '{http://schemas.microsoft.com/sharepoint/soap/}' + parameter)
        if value:
            sub.text = value

    # UpdateListItems Method
    def add_actions(self, data, kind):
        if not self.updates:
            updates = etree.SubElement(self.command, '{http://schemas.microsoft.com/sharepoint/soap/}updates')
            self.batch = etree.SubElement(updates, 'Batch')
            self.batch.set('OnError', 'Return')
            self.batch.set('ListVersion', '1')

        for index, row in enumerate(data, 1):
            method = etree.SubElement(self.batch, 'Method')
            method.set('ID', str(index))
            method.set('Cmd', kind)
            for key, value in row.items():
                field = etree.SubElement(method, 'Field')
                field.set('Name', key)
                field.text = str(value)

    # GetListFields Method
    def add_view_fields(self, fields):
        viewFields = etree.SubElement(self.command, '{http://schemas.microsoft.com/sharepoint/soap/}viewFields')
        viewFields.set('ViewFieldsOnly', 'true')
        ViewFields = etree.SubElement(viewFields, 'ViewFields')
        for field in fields:
            view_field = etree.SubElement(ViewFields, 'FieldRef')
            view_field.set('Name', field)

    # GetListItems Method
    def add_query(self, pyquery):
        query = etree.SubElement(self.command, '{http://schemas.microsoft.com/sharepoint/soap/}query')
        Query = etree.SubElement(query, 'Query')

        if 'Where' in pyquery:
            Query.append(pyquery['Where'])


    # UTF-8 encoding
    def __repr__(self):
        return (self.start_str + etree.tostring(self.envelope)).decode('utf-8')

    def __str__(self, pretty_print=False):
        return (self.start_str + etree.tostring(self.envelope, pretty_print=True)).decode('utf-8')
