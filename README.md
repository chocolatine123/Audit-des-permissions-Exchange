# Audit des permissions déléguées par les utilisateurs sous Exchange

Prérequis :

Nécessite Microsoft Exchange Web Services Managed API 1.2 ou versions supérieures
Lien de téléchargement : https://www.microsoft.com/en-us/download/details.aspx?id=42951

Nécessite les Outils d’administration de serveur distant pour Windows 10
Lien de téléchargement : https://www.microsoft.com/fr-FR/download/details.aspx?id=45520

Nécessite d'être membre de Organization Management

Nécessite en outre de posséder le rôle ApplicationImpersonation (présent dans le groupe Hygiene Management)

Nécessite d'être exécuté à partir d'une machine membre du domaine

OutputDir doit se trouver dans le même répertoire que ce script, ainsi que Get-MailboxFolderPermissionEWS.ps1
Ce dernier est à télécharger via ce lien : https://gallery.technet.microsoft.com/scriptcenter/Get-MailboxFolderPermission-fdf1f90f

Fonctionement

Ce script récupère les informations sur les administrateurs Exchange ainsi que sur les différentes permissions. Etant donné qu’un administrateur est ou peut être administrateur Exchange, la première partie de ce script sera de recenser les différents membres de ‘Administrators’, ‘Domain Admins’ et de ‘Enterprise Admins’, ainsi que ceux de ‘Organization Managment’, correspondant aux administrateurs Exchange à proprement parler. Le tout est stocké dans un fichier texte nommé AdminExchange<date_du_jour>.txt qui se trouve dans un répertoire réseau saisi au préalable par l’utilisateur. Ensuite, nous recensons la totalité des droits Exchange, enregistré dans le fichier DroitsExchange<date_du_jour>.txt qui se trouve dans ce même répertoire réseau. Cela étant pour ce qu’on appellera les ‘droits Admins’, c’est-à-dire qu’ils sont attribués par les administrateurs Exchange et stockés dans le Security Descriptor de l’utilisateur concerné.

Venons-en à présent aux droits qu’un utilisateur délègue à un autre, comme par exemple la lecture de son calendrier ou l’accès total à sa boite de réception, que l’on appellera ‘Droits Utilisateurs’. Bien évidemment, un simple utilisateur n’a pas accès à son Security Descriptor et ne peut par conséquent pas le modifier. Il faut savoir qu’une boite aux lettres est composée de plusieurs dossier (Boite de réception, d’envoi, dossier public, calendrier...) et dans le cas de délégation de droits par un utilisateur, une ACL est créée sur le dossier en question. S’il n’existe aucune commande pour vérifier ces droits, Microsoft a publié dernièrement un script dont on se servira pour les récupérer. Cependant, un membre d’Organization Managment’ n’a accès qu’au haut de la banque d’information. Il faudra de plus posséder le rôle Application Impersonation, présent dans le groupe Hygiène Management pour pouvoir exécuter ce script.

Les autres prérequis sont les suivants : 

- Télécharger le script Get-MailboxFolderPermissionEWS.ps1 qui nécessite Microsoft Exchange Web Services Managed API 1.2 ou toute version supérieure.
-	Installer les Outils d’administration de serveur distant pour Windows 10

Enfin, l’image ci-jointe montre l’exécution de cet outil, avec en haut à gauche l’exécution du script, en bas le répertoire de sortie et le rendu final en HTML à droite.

