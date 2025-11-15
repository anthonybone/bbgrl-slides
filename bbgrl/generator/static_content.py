def get_static_devotional_content():
    """Static devotional content extracted from the original generator.

    Keeping this separate prevents noise in the main generator and makes updates
    simple when text changes are needed.
    """
    return {
        "sacred_heart_hymns": [
            {"content": ""},
            {
                "content": (
                    "Heart of Jesus meek and mild. Hear oh hear thy feeble child "
                    "When the tempest's most severe. Heart of Jesus, hear. "
                    "Sweetly, we'll rest on thy Sacred Heart. Never from Thee. "
                    "O let us part!"
                )
            },
            {
                "content": (
                    "Hear then, Thy loving children's pray'r O Heart of Jesus, "
                    "Heart of Jesus hear."
                )
            },
            {"content": ""},
            {
                "content": (
                    "Oh Sacred Heart, Oh love divine. Do keep us near to Thee. "
                    "And make our love so like to Thine That we may holy be. "
                    "Heart of Jesus hear. Oh heart of love divine. Listen to our pray'r."
                )
            },
            {
                "content": (
                    "Make us always Thine. Oh temple pure, Oh house of gold. Our heaven here below. "
                    "What gifts unfurled, what wealth untold. From Thee do ever flow. Heart of Jesus hear. "
                    "Oh Heart of love divine. Listen to our pray'r. Make us always Thine."
                )
            },
        ],
        "post_communion_prayers": [
            {"content": ""},
            {
                "content": (
                    "Soul of Christ, make me holy. Body of Christ, save me. Blood of Christ, inebriate me. "
                    "Water from the side of Christ, wash me. Passion of Christ, make me strong. "
                    "O good Jesus, hear me. Hide me within your wounds."
                )
            },
            {
                "content": (
                    "Let me never be separated from You. Deliver me from the wicked enemy, "
                    "Call me at the hour of my death. And tell me to come to you that with Your saints "
                    "I may praise You forever. Amen."
                )
            },
            {
                "title": "PRAYER OF THANKSGIVING:",
                "content": (
                    "Lord God, I thank you through the Sacred Heart of Jesus, who is pleased to offer You "
                    "on our behalf continuous thanksgiving in the Eucharist."
                ),
            },
        ],
        "jubilee_prayer": [
            {"title": "THE JUBILEE PRAYER"},
            {
                "content": (
                    "Father in heaven, may the faith you have given us in your son, Jesus Christ, our brother, "
                    "and the flame of charity"
                )
            },
            {
                "content": (
                    "enkindled in our hearts by the Holy Spirit, reawaken in us the blessed hope for the coming "
                    "of your Kingdom."
                )
            },
        ],
        "st_joseph_prayer": [
            {
                "content": (
                    "To you, O blessed Joseph, do we come in our tribulation, and having implored the help of "
                    "your most holy Spouse, we confidently invoke your patronage also."
                )
            },
        ],
    }
