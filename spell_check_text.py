import spellchecker as sc
import sys

if __name__ == "__main__":
    sc = sc.SpellChecker()

    kwargs = {}
    for arg in sys.argv[1:]:
        if arg == '--no-auto':
            kwargs['auto'] = False
        if arg == '--debug':
            kwargs['debug'] = True
        if arg == '--no-suggestions':
            kwargs['suggest'] = False
        if arg == '--no-google':
            kwargs['google_sc'] = False
    
    sc.spell_check_text(**kwargs)
    