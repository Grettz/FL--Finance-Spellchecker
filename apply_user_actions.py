import spellchecker as sc
import time

if __name__ == "__main__":
    sc = sc.SpellChecker(_spacy=False)
    
    start_time = time.time()
    sc.apply_user_actions()
    time_taken = time.time() - start_time
    print(f'\nUser changes processed in {time_taken:.3f}s')
    