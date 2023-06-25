from PhilippinesRegions import PhilippinesRegions

philippines = PhilippinesRegions()

selected_region = "Caraga"
selected_provinces = philippines.get_provinces(selected_region)

print(f"\nProvinces in {selected_region}:")
if selected_provinces:
    for province in selected_provinces:
        print(province)
else:
    print("Region not found or has no provinces.")
