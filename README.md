# Excel-Functions
Excel VBA functions to help you deals with IP address resolution, DNS, FQDN and other network stuffs

The .txt files are the raw VBA code. The code of the functions in these files must be added in a module using the developper tools in Excel.

getDomain(URL) return the FQDN (Fully qualified domain name) from a string containing a host, subdomain and domain.

getIP(FQDN) does a NSLokkup on a hostname and return the associated IP address

getNameFromIP(IPAddress) does a NSLookup on an IP address and return the associated FQDN

isIPv4(String) return TRUE if the string is a valid IPv4 address (FALSE otherwise)

IP2BIN(IPAddress) convert an IPv4 address string into a string with its 32 bits binary equivalent

inSubnet(IPAddress, Subnet address, Subnet_CIDR) returns IN if the provided IP address belongs to the subnet and OUT if it is not part of it.

GetMatrixValue(Martrix range, row name, column name) returns the value of a specific cell in a matrix based on the value of the headers of the row and column of the matrix. It is quite useful to check if a flow from one network to another network is authorized or not just by using a matrix with all networks'name and the default autorisation.

GetSubnetName(IPaddress, List of subnets) returns the name of the subnet to wich IPaddress belongs based on a list of subnets, CIDR and names.

