# application-mise-jour-des-prix-VBA-

Pour répondre aux évolutions tarifaires, Mon client avait besoin de mettre à jour rapidement les prix de ses produits et services sur plusieurs centaines de fichiers Excel, un travail manuel jusqu'alors très chronophage et propice aux erreurs.

La personne chargée de ce travail doit travailler manuellement sur des centaines de fichiers et chaque fichier contient des centaines de lignes, Ce travail est très fatigant et fastidieux, il prend beaucoup de temps et il peut être une source d’erreur aussi.

![image](https://github.com/user-attachments/assets/e0b374b1-4bae-4b9d-8136-a57b05186db5)
![image](https://github.com/user-attachments/assets/c32f2871-ddd1-4733-b9c5-e12e6cea9a49)

Pour résoudre ce problème, il a fallu développer un programme qui modifie les prix de chaque produit de manière fluide et sans erreur et automatique.

# Application Mise à jour de prix version 1
Dans la première version du programme, l’idée était de travailler sur chaque fichier dont on voulait modifier les prix séparément, chaque fichier contenant plusieurs Articles et avoir un seul type de prix (grilles) A, B et C, qu’on peut remarquer sur le nom de chaque fichier.
![image](https://github.com/user-attachments/assets/8d087bcc-16dd-4c02-8309-6df05875f3b4)

On a utilisé deux autre fichier Excel, le premier fichier Excel contient les anciens prix « 00 - Produits - Anciens prix ABC.xls » et le deuxième contient les nouveaux prix « 00 - Produits - Nouveaux prix ABC.xls ».
![image](https://github.com/user-attachments/assets/f462a6c9-8819-483a-8e1f-be8a1d4bfb06)

L’opération était de vérifier le prix de chaque article qui est dans le fichier de grille, s’il a le même prix de « Anciens prix » ou non, si oui on doit prendre le prix qui est dans le « Nouveaux prix » si non on vérifier est ce qu‘il est le nouveau prix ou non.

Pour chaque grille on utilise une formule différente car chaque grille correspondant à une colonne dans le fichier « 00 – Produits - Nouveaux prix ABC.xls » comme indiqué dans la figure
22. Donc j’ai créé trois formulaires qui on peut les utiliser pour tous les types de grilles

La deuxième tâche de ce programme consiste à filtrer les articles dont le prix n'a pas changé et à les placer sur une nouvelle feuille qui s’appelle « Ajout » afin de modifier leur prix.
La troisième tâche qui doit être effectuée par ce programme est d’exporter le fichier traité vers un dossier.
Afin de mettre en œuvre ces deux tâches précédentes, j'ai dû trouver des solutions en utilisant le langage de programmation VBA.
Grâce à une analyse des besoins qui j’ai faite avec l’aide de mon tuteur j’ai décidé de développer un programme avec 4 boutons et une barre de choix de grille (A, B et C) :
•	Importer : L'import des fichiers sur lesquels nous voulons travailler et affiche le nom du fichier dans une fenêtres.
•	Prix : barre de choix de grille (A, B et C) qui dépend du nom de fichier
•	Mise prix : Contient toutes les opérations nécessaires qui doivent être appliquées à chaque grille et chaque  article pour avoir à la fin deux onglets bien traités (Modification et Ajout).
•	Exporter : L’export des fichiers traités dans le même dossier d’import avec le même nom du fichier.
•	Supprimer : supprime tous les onglets et tout le contenu dans le fichier de programme, recommandés après chaque import.
Chaque bouton doit avoir une gestion d’erreur qui affiche un message en cas de mauvaise sélection du bouton ou en cas de fin d’exécution.
Les fichiers nécessaires pour exécuter ce programme :
-	00 - Produits - Anciens prix ABC.xls
-	00 - Produits - Nouveaux prix ABC.xls

 
      Application de mise à jour de prix version 1

  ![image](https://github.com/user-attachments/assets/48ff7f87-bd76-41f5-bbb0-b5a57e30b6ea)

# Application Mise à jour de prix version 2

Après avoir créé la première version qui traite grille par grille, nous avons remarqué que nous devions importer le fichier à chaque fois, ce programme fonctionne parfaitement si nous n'avons pas beaucoup de fichiers.
C'était une solution séquentielle en première période, mais au bout d'un moment nous avons dû travailler sur un seul fichier qui regroupe tous les fichiers et toutes les grilles pour résoudre à chaque fois le problème d'import.
Nous avons dû créer un programme de traitement automatique sur plusieurs grilles, le programme doit être capable de reconnaitre chaque grille de chaque article et de travailler dessus.

![image](https://github.com/user-attachments/assets/7c53b814-06c6-437e-b4c2-a9d611ccf438)


Il a fallu trouver une solution pour travailler sur un fichier contenant des milliers de lignes et trois grilles mélangées dans l'ensemble du fichier.

![image](https://github.com/user-attachments/assets/28886abc-efd4-4db0-b86f-7a54f980f648)
Exemple : Nombre de lignes « 119226 »

Travailler sur ce programme a été le plus difficile pendant mon année d’alternance, afin que je puisse trouver une solution à ce problème, j'ai fait une formation en ligne pour apprendre la programmation VBA avancée en anglais.
La solution consistait à utiliser plusieurs filtres qui filtrer les articles en fonction de grilles et les ‘enregistres dans des dictionnaires.
J’ai utilisé deux filtres le premier pour les grilles et deuxième pour les articles, chaque filtre enregistrant des informations non récursives dans un dictionnaire :
![image](https://github.com/user-attachments/assets/6e2e4e41-9c96-4012-b02c-fe91fec111f3)
Exemple de code de filtre et dictionnaires

Avec les valeurs enregistrées dans les dictionnaires, on peut travailler sur chaque article séparément et modifier les tarifs en fonction de chaque grille.
Le programme conserve toujours les mêmes caractéristiques et les résultats souhaités avec beaucoup moins de temps que la première version.
Après des améliorations sur le code pour faire la mise à jour automatique de plusieurs grilles et plusieurs fonctionnalités sont ajoutées, dont le compteur de lignes et le compteur de temps qui apparaît à la fin du processus, comme indiqué sur les images :

![image](https://github.com/user-attachments/assets/b0a8352a-7aa4-4bbf-aca6-09b589848d3a)
Nombre de lignes importées


![image](https://github.com/user-attachments/assets/e63acf4e-9ffe-44e8-b10b-404d31e26b2e)
Durée de traitement


![image](https://github.com/user-attachments/assets/38b3d052-fe68-43fc-9704-c1d588d66993)

