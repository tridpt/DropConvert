import sys


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        from dropconvert.smoke import main as smoke_main

        raise SystemExit(smoke_main(sys.argv[2]))
    from dropconvert.ui import main

    raise SystemExit(main())
