![http://img232.imageshack.us/img232/922/monorunneruu0.png](http://img232.imageshack.us/img232/922/monorunneruu0.png)

# MONORunner: Che cos'è ... #

Se avete già utilizzato un'applicazione MONO su piattaforma MS Windows vi sarete accorti che per avviare l'applicazione dovete indicare il path dell'eseguibile di MONO (mono.exe) seguito dal path dell'applicazione.
L'eseguibile di **MONORunner** si occuperà di lanciare correttamente l'applicazione per MONO andandosi a cercare il giusto path dell'eseguibile di MONO (mono.exe).

# MONORunner: Come si usa ... #

Posso usare **MONORunner** per creare pacchetti d'installazione su piattaforma MS Windows per applicazione MONO.
Esempio:
Supponiamo di avere un'applicazione MONO dal nome _myApplication.exe_ e vogliamo distribuirla. Possiamo creare un setup che contenga sia _myApplication.exe_ che _monorunner.exe_ ma i link per avviare l'applicazione dovranno puntare a _monorunner.exe_ e come parametro dovranno avere _myApplication.exe_ in quanto **MONORunner** una volta intercettata la posizione corretta dell'eseguibile mono.exe, utilizza proprio il parametro passato sulla riga di comando per avviare l'applicazione.

[![](http://img505.imageshack.us/img505/5232/monopoweredbm0.png)](http://www.mono-project.com)