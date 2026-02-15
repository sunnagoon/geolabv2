from app.db import init_db
from app.ui.app import GeoLabApp


def main():
    init_db()
    app = GeoLabApp()
    app.mainloop()


if __name__ == "__main__":
    main()