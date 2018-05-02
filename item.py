class item
    __ShopType = None
    __ItemName = None
    __Price = None
    __Level = None
    __Skill-1 = None
    __Skill-2 = None
    __Skill-3 = None
    __Rarity = None
    __Description = None

    def set_ShopType(self, ShopType):
        self.__ShopType = ShopType

    def get_ShopType(self):
        return self.__ShopType

    def set_ItemName(self, ItemName):
        self.__ItemName = ItemName

    def get_ItemName(self):
        return self.__ItemName

    def set_Price(self, Price):
        self.__Price = Price

    def get_Raw_Price(self):
        return self.__Price

    def get_style_Price(self):
        # This is where the currency convert function goes
        return 
