import hml_equation_parser as hp

string = hp.eq2latex("LEFT ⌊ a+b RIGHT ⌋")
string2 = hp.eq2latex("THEREFORE  TRIANGLE  rm PAC` == ` TRIANGLE  rm PBE")
string3 = hp.eq2latex("f LEFT ( x RIGHT ) `=` {k} over {x-2} +k")
print(string3)

