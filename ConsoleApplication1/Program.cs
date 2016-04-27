using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;

namespace AD
{
    class Program
    {
        static string ip;
        static string loginAD;
        static string passwordAD;

        static void Main(string[] args)
        {
            /* http://morpheus.developpez.com/addotnet/ADCSharp/ */

            Connection();

            Console.ReadLine();
        }


        // Connection au serveur Active Directory
        // Manque test de la connection
        static void Connection()
        {
            Console.Write("Veuillez entrer l'adresse IP de votre serveur Active Directory : ");
            ip = Console.ReadLine();
            Console.WriteLine();

            Console.Write("Veuillez entrer votre nom d'utilisateur : ");
            loginAD = Console.ReadLine();
            Console.WriteLine();

            Console.Write("Veuillez entrer votre mot de passe : ");
            passwordAD = MaskPassword();
            Console.WriteLine();

            DirectoryEntry ldap = new DirectoryEntry("LDAP://" + ip, loginAD, passwordAD);

            try
            {
                DirectorySearcher searcher = new DirectorySearcher(ldap);
                searcher.Filter = "(objectClass=user)";
                searcher.FindAll();
            }

            catch (Exception Ex)
            {
                Console.WriteLine();
                Console.WriteLine(Ex.Message + "Veuillez réessayer.");
                Console.WriteLine();
                Connection();
            }

            Menu(ldap);
        }

        // Méthode pour masquer le mot de passe
        static string MaskPassword()
        {
            /* https://social.msdn.microsoft.com/Forums/fr-FR/55e423d6-7917-4e7d-822d-ce1adcd547c6/comment-masquer-les-caractres-de-mot-de-passe-dans-la-console-lors-de-la-saisie-?forum=visualcsharpfr */

            Stack<char> stack = new Stack<char>();
            ConsoleKeyInfo consoleKeyInfo;

            while ((consoleKeyInfo = Console.ReadKey(true)).Key != ConsoleKey.Enter)
            {
                if (consoleKeyInfo.Key != ConsoleKey.Backspace)
                {
                    stack.Push(consoleKeyInfo.KeyChar);
                    Console.Write("*");
                }

                else
                {
                    Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    Console.Write(" ");
                    Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    if (stack.Count > 0) stack.Pop();
                }
            }

            return stack.Reverse().Aggregate(string.Empty, (pass, kc) => pass += kc.ToString());
        }

        // Menu principal (fini)
        static void Menu(DirectoryEntry ldap)
        {
            Console.Clear();

            Console.WriteLine("Que souhaitez-vous faire ?");
            Console.WriteLine();
            Console.WriteLine("1 - Créer un nouvel utilisateur");
            Console.WriteLine("2 - Voir la liste des utilisateurs");
            Console.WriteLine("3 - Voir un utilisateur en détail");
            Console.WriteLine("4 - Modifier un utilisateur");
            Console.WriteLine("5 - Supprimer un utilisateur");
            Console.WriteLine("6 - Quitter");
            Console.WriteLine();
            Console.Write("Appuyer sur la touche correspondant à votre choix : ");

            string choix = Console.ReadLine();

            switch (choix)
            {
                case "1":
                    Console.Clear();
                    Create(ldap);
                    break;

                case "2":
                    Console.Clear();
                    Read(ldap);
                    break;

                case "3":
                    Console.Clear();
                    ReadDetails(ldap);
                    break;

                case "4":
                    Console.Clear();
                    Update(ldap);
                    break;

                case "5":
                    Console.Clear();
                    Delete(ldap);
                    break;

                case "6":
                    Exit(ldap);
                    break;
            }
        }

        // Créer un utilisateur (fini)
        static void Create(DirectoryEntry ldap)
        {
            Console.Write("Veuillez entrer le nom de l'utilisateur : ");
            string nom = Console.ReadLine();
            Console.WriteLine();

            Console.Write("Veuillez entrer le prénom de l'utilisateur : ");
            string prenom = Console.ReadLine();
            Console.WriteLine();

            Console.Write("Veuillez entrer le login de l'utilisateur : ");
            string loginUser = Console.ReadLine();
            Console.WriteLine();

            // Création de l'utilisateur et initialisation de ses propriétés
            DirectoryEntry user = ldap.Children.Add("cn= " + prenom + " " + nom, "user");
            user.Properties["sn"].Add(nom);
            user.Properties["givenName"].Add(prenom);
            user.Properties["SAMAccountName"].Add(loginUser);
            user.Properties["displayname"].Add(prenom + " " + nom);
            user.Properties["name"].Add(prenom + " " + nom);

            bool loop = true;

            while (loop == true)
            {
                Console.Write("Souhaitez-vous entrer une description ? (O/N) : ");
                string answer = Console.ReadLine();
                answer = answer.ToUpper();

                Console.WriteLine();

                if (answer == "O" || answer == "OUI")
                {
                    Console.WriteLine("Veuillez entrer une courte description : ");
                    string description = Console.ReadLine();
                    user.Properties["description"].Add(description);
                    Console.WriteLine();
                    loop = false;
                }

                if (answer == "N" || answer == "NON")
                {
                    loop = false;
                }

                else
                {
                    Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }

            while (loop == false)
            {
                Console.Write("Souhaitez-vous renseigner une adresse mail ? (O/N) :");
                string answer = Console.ReadLine();
                answer = answer.ToUpper();

                Console.WriteLine();

                if (answer == "O" || answer == "OUI")
                {
                    Console.WriteLine("Veuillez entrer l'adresse mail de l'utilisateur : ");
                    string mail = Console.ReadLine();
                    user.Properties["mail"].Add(mail);
                    Console.WriteLine();
                    loop = true;
                }

                if (answer == "N" || answer == "NON")
                {
                    loop = true;
                }

                else
                {
                    Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }

            while (loop == true)
            {
                Console.Write("Souhaitez-vous renseigner un numéros de téléphone ? (O/N) :");
                string answer = Console.ReadLine();
                answer = answer.ToUpper();

                Console.WriteLine();

                if (answer == "O" || answer == "OUI")
                {
                    Console.WriteLine("Veuillez entrer le numéros de téléphone de l'utilisateur : ");
                    string numTel = Console.ReadLine();
                    user.Properties["TelephoneNumber"].Add(numTel);
                    Console.WriteLine();
                    loop = false;
                }

                if (answer == "N" || answer == "NON")
                {
                    loop = false;
                }

                else
                {
                    Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }

            while (loop == false)
            { 
                Console.Write("Souhaitez-vous assigner une machine à cet utilisateur ? (O/N) :");
                string answer = Console.ReadLine();
                answer = answer.ToUpper();

                Console.WriteLine();

                if (answer == "O" || answer == "OUI")
                {
                    Console.WriteLine("Veuillez entrer le nom de la machine : ");
                    string machine = Console.ReadLine();
                    user.Properties["userworkstations"].Add(machine);
                    Console.WriteLine();
                    loop = true;
                }

                if (answer == "N" || answer == "NON")
                {
                    loop = true;
                }

                else
                {
                    Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }

            while (loop == true)
            {
                Console.Write("Souhaitez-vous ajouter une entreprise ? (O/N) : ");
                string answer = Console.ReadLine();
                answer = answer.ToUpper();

                Console.WriteLine();

                if (answer == "O" || answer == "OUI")
                {
                    Console.WriteLine("Veuillez entrer le nom de l'entreprise : ");
                    string entreprise = Console.ReadLine();
                    user.Properties["company"].Add(entreprise);
                    Console.WriteLine();
                    loop = false;
                }

                if (answer == "N" || answer == "NON")
                {
                    loop = false;
                }

                else
                {
                    Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }

            // On envoie les modifications
            user.CommitChanges();

            // Définition du mot de passe. 
            while (loop == false)
            {
                Console.Write("Veuillez entrer le mot de passe de l'utilisateur : ");
                string passwordUser = MaskPassword();
                Console.WriteLine();

                Console.Write("Veuillez confirmer le mot de passe : ");
                string passwordUserConfirmation = MaskPassword();
                Console.WriteLine();

                if (passwordUser == passwordUserConfirmation)
                {
                    user.Invoke("SetPassword", new object[] { passwordUser });
                    Console.WriteLine("Le mot de passe a été créé avec succès !");
                    Console.WriteLine();
                    loop = true;
                }

                else
                {
                    Console.WriteLine("Les mots de passe ne sont pas identique. Veuillez réessayer.");
                    Console.WriteLine();
                }
            }
            

            // Activation du compte : 
            user.Properties["userAccountControl"].Value = 0x0200;

            // On réenvoie les modifications 
            user.CommitChanges();

            Console.WriteLine("L'utilisateur a été correctement enregistré.");
            Console.WriteLine("Appuyer sur n'importe quelle touche pour revenir au menu principal.");
            Console.ReadLine();
            Menu(ldap);
        }

        // Lister tous les utilisateurs de l'AD (Fini)
        static void Read(DirectoryEntry ldap)
        {
            DirectorySearcher searcher = new DirectorySearcher(ldap);
            searcher.Filter = "(objectClass=user)";

            foreach(SearchResult result in searcher.FindAll())
            {
                DirectoryEntry user = result.GetDirectoryEntry();
                Console.WriteLine("Login  : " + user.Properties["SAMAccountName"].Value);
                Console.WriteLine("Nom    : " + user.Properties["sn"].Value);
                Console.WriteLine("Prénom : " + user.Properties["givenName"].Value);
                Console.WriteLine("Email  : " + user.Properties["mail"].Value);
                Console.WriteLine("Tél    : " + user.Properties["TelephoneNumber"].Value);
                Console.WriteLine("Description : " + user.Properties["description"].Value);
                Console.WriteLine("----------------------------");
            }

            Console.WriteLine();
            Console.WriteLine("Appuyer sur n'importe quelle touche pour revenir au menu principal.");
            Console.ReadLine();
            Menu(ldap);
        }

        // Voir un utilisateur en détail (fini)
        static void ReadDetails(DirectoryEntry ldap)
        {
            DirectorySearcher searcher = new DirectorySearcher(ldap);

            // Recherche par login car aucun utilisateur ne peut avoir le même
            Console.Write("Veuillez entrer le login de l'utilisateur : ");
            string loginUser = Console.ReadLine();
            Console.WriteLine();

            searcher.Filter = "(SAMAccountName=" + loginUser + ")";
            SearchResult result = searcher.FindOne();

            if (result != null)
            {
                // L'utilisateur existe, on liste les champs
                ResultPropertyCollection fields = result.Properties;
                foreach (String ldapField in fields.PropertyNames)
                {
                    // Il peut y avoir plusieurs objets dans chaque champs (ex: appartenance à des groupes) 
                    foreach (Object myCollection in fields[ldapField])
                    {
                        Console.WriteLine(String.Format("{0,-20} : {1}", ldapField, myCollection.ToString()));
                    }
                }

            }

            else
            {
                // L'utilisateur n’existe pas  
                Console.WriteLine("Utilisateur non trouvé !");

                while (true)
                { 
                    Console.Write("Souhaitez-vous réessayer ? (O/N) :");
                    string answer = Console.ReadLine();
                    answer = answer.ToUpper();

                    if (answer == "O" || answer == "OUI")
                    {
                        ReadDetails(ldap);
                    }

                    if (answer == "N" || answer == "NON")
                    {
                        Menu(ldap);
                    }

                    else
                    {
                        Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                        Console.WriteLine();
                    }
                }
            }

            Console.WriteLine();
            Console.WriteLine("Appuyer sur n'importe quelle touche pour revenir au menu principal.");
            Console.ReadLine();
            Menu(ldap);
        }

        // Modifier un utilisateur (fini)
        static void Update(DirectoryEntry ldap)
        {
            DirectorySearcher searcher = new DirectorySearcher(ldap);

            // Recherche par login car aucun utilisateur ne peut avoir le même
            Console.Write("Veuillez entrer le login de l'utilisateur que vous souhaitez modifier : ");
            string loginUser = Console.ReadLine();
            Console.WriteLine();

            searcher.Filter= "(SAMAccountName=" + loginUser + ")";
            SearchResult result = searcher.FindOne();

            if(result != null)
            {
                // On récupère l'objet 
                DirectoryEntry user = result.GetDirectoryEntry();

                Console.WriteLine("Que souhaitez-vous modifier ?");
                Console.WriteLine();
                Console.WriteLine("1  - Modifier le nom");
                Console.WriteLine("2  - Modifier le prénom");
                Console.WriteLine("3  - Modifier le login");
                Console.WriteLine("4  - Modifier la description");
                Console.WriteLine("5  - Modifier l'email");
                Console.WriteLine("6  - Modifier le numéros de téléphone");
                Console.WriteLine("7  - Modifier la machine assignée");
                Console.WriteLine("8  - Modifier l'entreprise");
                Console.WriteLine("9  - Modifier le mot de passe");
                Console.WriteLine("10 - Retour au menu principal");
                Console.WriteLine();
                Console.Write("Veuillez entrer le numéro correspondant à votre choix : ");

                string Choix = Console.ReadLine();

                switch (Choix)
                {
                    #region Modification Nom
                    case "1":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer le nouveau nom de l'utilisateur : ");
                        string Nom = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["sn"].Value = Nom;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("Le nom a été changé avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Prenom
                    case "2":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer le nouveau prénom de l'utilisateur : ");
                        string Prenom = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["givenName"].Value = Prenom;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("Le prénom a été changé avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Login
                    case "3":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer le nouveau login de l'utilisateur : ");
                        string Login = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["SAMAccountName"].Value = Login;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("Le login a été changé avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Description
                    case "4":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer la nouvelle description de l'utilisateur : ");
                        string Description = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["description"].Value = Description;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("La description a été changée avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Mail
                    case "5":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer la nouvelle adresse mail de l'utilisateur : ");
                        string Mail = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["mail"].Value = Mail;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("L'adresse mail a été changée avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Telephone
                    case "6":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer le nouveau numéro de téléphone de l'utilisateur : ");
                        string NumTel = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["TelephoneNumber"].Value = NumTel;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("Le numéro de téléhpone a été changé avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Machine
                    case "7":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer le nom de la machine asignée à l'utilisateur : ");
                        string Machine = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["userworkstations"].Value = Machine;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("La machine a été assignée avec succès !");
                        Console.WriteLine();
                        Update(ldap);
                        break;
                    #endregion

                    #region Modification Entreprise
                    case "8":
                        Console.WriteLine();
                        Console.Write("Veuillez entrer l'entreprise de l'utilisateur : ");
                        string Company = Console.ReadLine();
                        Console.WriteLine();

                        user.Properties["Company"].Value = Company;

                        // On envoie les changements
                        user.CommitChanges();

                        Console.WriteLine("L'entreprise a été changée avec succès !");
                        Console.WriteLine();
                        Update(ldap);

                        break;
                    #endregion

                    #region Modification Password
                    case "9":
                        Console.WriteLine();
                        while (true)
                        { 
                            Console.Write("Veuillez entrer le nouveau mot de passe de l'utilisateur : ");
                            string PasswordUser = MaskPassword();
                            Console.WriteLine();

                            Console.Write("Veuillez confirmer le mot de passe : ");
                            string PasswordUserConfirmation = MaskPassword();
                            Console.WriteLine();

                            if (PasswordUser == PasswordUserConfirmation)
                            {
                                user.Invoke("SetPassword", new object[] { PasswordUser });
                                Console.WriteLine("Le mot de passe a été modifié avec succès !");
                                Update(ldap);
                            }

                            else
                            {
                                bool loop = true;

                                while(loop == true)
                                { 
                                    Console.WriteLine("Les mots de passe ne sont pas identique.");
                                    Console.WriteLine();
                                    Console.Write("Voulez-vous réessayer ? (O/N) :");
                                    string Answer = Console.ReadLine();
                                    Answer = Answer.ToUpper();

                                    if (Answer == "O" || Answer == "OUI")
                                    {
                                        loop = false;
                                    }

                                    if (Answer == "N" || Answer == "NON")
                                    {
                                        Update(ldap);
                                    }

                                    else
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                                        Console.WriteLine();
                                    }
                                }
                            }
                        }
                        break;
                    #endregion

                    case "10":
                        Menu(ldap);
                        break;
                }
            }

            else
            {  
                // L'utilisateur n’existe pas  
                Console.WriteLine("Utilisateur non trouvé!");

                while (true)
                {
                    Console.Write("Souhaitez-vous réessayer ? (O/N) :");
                    string answer = Console.ReadLine();
                    answer = answer.ToUpper();

                    if (answer == "O" || answer == "OUI")
                    {
                        Update(ldap);
                    }

                    if (answer == "N" || answer == "NON")
                    {
                        Menu(ldap);
                    }

                    else
                    {
                        Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                        Console.WriteLine();
                    }
                }
            }
        }

        // Supprimer un utilisateur (fini)
        static void Delete(DirectoryEntry ldap)
        {
            DirectorySearcher searcher = new DirectorySearcher(ldap);

            // Recherche par login car aucun utilisateur ne peut avoir le même
            Console.Write("Veuillez entrer le login de l'utilisateur : ");
            string loginUser = Console.ReadLine();
            Console.WriteLine();

            searcher.Filter = "(SAMAccountName=" + loginUser + ")";
            SearchResult result = searcher.FindOne();

            if (result != null)
            {
                while(true)
                { 
                    // L'utilisateur existe, on liste les champs
                    Console.WriteLine("Etes-vous sûr de vouloir supprimer l'utilisateur ? (O/N) : ");
                    string answer = Console.ReadLine();
                    answer = answer.ToUpper();

                    if (answer == "O" || answer == "OUI")
                    {
                        // On récupère l'objet 
                        DirectoryEntry user = result.GetDirectoryEntry();
                        user.DeleteTree();

                        Console.WriteLine("L'utilisateur a été supprimé avec succès !");
                        Console.WriteLine("Appuyer sur n'importe quelle touche pour revenir au menu principal");
                        Menu(ldap);
                    }

                    if (answer == "N" || answer == "NON")
                    {
                        Menu(ldap);
                    }

                    else
                    {
                        Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                        Console.WriteLine();
                    }
                }

            }

            else
            {
                // L'utilisateur n’existe pas  
                Console.WriteLine("Utilisateur non trouvé !");

                while (true)
                {
                    Console.Write("Souhaitez-vous réessayer ? (O/N) :");
                    string answer = Console.ReadLine();
                    answer = answer.ToUpper();

                    if (answer == "O" || answer == "OUI")
                    {
                        Delete(ldap);
                    }

                    if (answer == "N" || answer == "NON")
                    {
                        Menu(ldap);
                    }

                    else
                    {
                        Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                        Console.WriteLine();
                    }
                }
            }
        }

        // Quitter correctement l'application (fini)
        static void Exit(DirectoryEntry ldap)
        {
            Console.Write("Voulez-vous vraiment quitter ? (O/N) : ");
            string answer = Console.ReadLine();
            answer = answer.ToUpper();

            if (answer == "O" || answer == "OUI")
            {
                ldap.Close();
                Environment.Exit(0);
            }

            if (answer == "N" || answer == "NON")
            {
                Menu(ldap);
            }

            else
            {
                Console.WriteLine("Je n'ai pas compris votre réponse. Veuillez réessayer.");
                Console.WriteLine();
                Exit(ldap);
            }
        }
    }
}
