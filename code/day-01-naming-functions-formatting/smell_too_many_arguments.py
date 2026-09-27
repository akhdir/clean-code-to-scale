def create_shipment(name, street, city, state,
                     zip_code, country, weight_kg,
                     service_level):
    ...

# 8 positional args - the call site tells you nothing about which is which
create_shipment("J. Rivera", "12 Elm St", "Austin",
                "TX", "78701", "US", 4.2, "express")
