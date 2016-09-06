Excel-VBA-NetworkTools contains Excel VBA functions to help you deal with IP addresses resolution, DNS, FQDN and other network related stuffs with Excel

The code of the functions in this file must be added in a module using the developper tools in Excel.

- getDomain(URL as a String) return the FQDN (Fully qualified domain name) from a string containing a host, subdomain and domain.

- getIP(FQDN as a String) does a NSLokkup on a hostname and return the associated IP address

- getNameFromIP(IPAddress as a String) does a NSLookup on an IP address and return the associated FQDN

- isIPv4(IPaddress as a String) return a boolean set to TRUE if IPaddress is a valid IPv4 address (FALSE otherwise)

- IP2BIN(IPAddress as a String) convert an IPv4 address string into a string with its 32 bits binary equivalent

- bin2IP(IPaddress as a String) convert a binary format IPv4 address into a string with the decimal value of each 4 bytes

- inSubnet(IPAddress as a String, Subnet address as a String, Subnet_CIDR as an Integer) returns the string "IN" if the provided IP address belongs to the subnet and the string "OUT" if it is not part of it.

- GetSubnetLowIP(IPaddress as a String, CIDR as an Integer) return a string with the lowest IP address in the subnet's range (the network address)

- GetSubnetHighIP(IPaddress a a string, CIDR as an Integer) return a string with the highest IP address in the subnet's range (the broadcast address)

- GetSubnetIPRange(IPaddress a a string, CIDR as an Integer) return a string with the lowest and the highest IP address in the subnet's range separated by a dash (e.g. 10.0.0.0 - 10.255.255.255)

- GetMatrixValue(Martrix range, row name, column name) returns the value of a specific cell in a matrix based on the value of the headers of the row and column of the matrix. It is quite useful to check if a flow from one network to another network is authorized or not just by using a matrix with all networks'name and the default autorisation.

- GetSubnetName(IPaddress, List of subnets) returns the name of the subnet to wich IPaddress belongs based on a list of subnets, CIDR and names.

- howmanybit(dec as a Long) return the number of bits needed to represent the decimal number dec in a binary format

- dec2bin(dec as long) return a string with the value of dec expressed in binary format
