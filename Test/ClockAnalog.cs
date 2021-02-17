using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace Test
{
    public partial class ClockAnalog : UserControl
    {

        public ClockAnalog()
        {
            InitializeComponent();
        }
        // booleen pour savoir si on déplace déplacer la fenetre ou pas
        private bool bStartDrag = false;
        // position de la souris sur l'ecran
        private int mx, my;
        // precalcule de l'angle que represente 1 seconde (ou 1 minute), en radian
        private double uSec = 6 * Math.PI / 180;
        // precalcule de l'angle que represente 1 heure, en radian
        private double uHour = 30 * Math.PI / 180;
        // precalcul de PI/2
        private double HalfPi = Math.PI / 2;

        // precalcul du centre de l'image (en gros ca sert à rien de précalculer mais bon, j'etais parti dans un délire)
        private int CenterX, CenterY;


        private void ClockAnalog_Load(object sender, EventArgs e)
        {// MàJ des valeurs de départ
            mx = this.Top;
            my = this.Left;

            CenterX = pictureBox1.Width / 2;
            CenterY = pictureBox1.Height / 2;

            label1.Text = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" +
                          DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" +
                          DateTime.Now.Second.ToString().PadLeft(2, '0');

        }



		private void pictureBox1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			// Si on n'est pas en train de trainer la form
			// on met à jour toutes les infos et on bascule le flag StartDrag
			if (!bStartDrag)
			{
				mx = PointToScreen(new Point(e.X, e.Y)).X - this.Left;
				my = PointToScreen(new Point(e.X, e.Y)).Y - this.Top;
				bStartDrag = true;
			}

		}

		private void pictureBox1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			// Lorsqu'on relache le bouton de la souris
			// on reinit les variables
			bStartDrag = false;

		}

		private void label1_Click(object sender, EventArgs e)
		{

		}

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void pictureBox1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			// Si on est en train de cliquer sur la fenetre
			// on la déplace au prorata de l'endroit d'ou on saisit la fenetre
			// sans ca, la fenetre placerait son origine, directement sous le pointeur
			if (bStartDrag)
			{
				this.Left = PointToScreen(new Point(e.X, e.Y)).X - mx;
				this.Top = PointToScreen(new Point(e.X, e.Y)).Y - my;
			}
		}

		// la partie la fun : le dessin de l'horloge 
		private void timer2_Tick(object sender, System.EventArgs e)
		{

			double osx, osy, sx, sy;
			double omx, omy, mmx, mmy;
			double ohx, ohy, hx, hy;
			double curSec, curMin, curHour;

			// declaration de l'objet qui sert à dessiner
			System.Drawing.Pen myPen = new System.Drawing.Pen(System.Drawing.Color.GhostWhite);

			// declaration de l'objet qui sert à définir la brosse de remplissage
			System.Drawing.Brush myBrush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);

			// declaration de l'objet dans lequel on peut dessiner. Ici, l'objet corespond
			// au format de la picturebox
			System.Drawing.Graphics formGraphics = pictureBox1.CreateGraphics();

			// c'est pour que le dessin soit antialiasé.
			// Commentez la ligne suivante pour voir le rendu sans ça, c'est... édifiant !
			formGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

			// Seconde courante
			curSec = (double)DateTime.Now.Second;

			// Minute courante + l'équivalent en base 10 des secondes courantes
			// Cette petite astuce permet d'avoir un mouvement "smoothé" de l'aiguille
			// car plus précis que de minute en minute
			curMin = (double)DateTime.Now.Minute + (curSec / 60.0);

			// Heure courante, avec la même astuce qu'utilisée pour les minutes
			curHour = (double)DateTime.Now.Hour + (curMin / 60.0);

			/* Un peu de math !
			 sur le cercle trigo (rayon 1) : x = cos, y = sin
			 donc si on multiplie les coordonnées par le mm nombre N, on obtient
			 un cercle de rayon N (logique non ?!)
			 Exple : pour tracer point à point un cercle de rayon 10, on aurait (code faux car en degrés, mais plus facile à comprendre)
					for (int i=0; i<360; i++)
					{
						x = cos(i) * 10;
						y = sin(i) * 10;
						dessinepoint(x, y)
					}
			 Il suffit ensuite d'ajouter les coordonnées du centre du cercle pour effectuer
			 une simple translation de repère 
			 Pour résumer : 
						x = cos(angle)*rayon + origineX
						y = sin(angle)*rayon + origineY

			A partir de ces éléments, rien n'est plus facile de déterminer une position sur un cercle
			lorsqu'on connait l'angle
			*/

			// On a précalculé l'angle que représente une seconde, on le multiplie
			// donc par le nombre de secondes. On y retire PI/2 (90°) pour que
			// ca tourne dans le bon sens : Le sens trigo est inverse au sens des aiguilles !
			sx = Math.Cos((uSec * curSec) - HalfPi) * 50 + (double)CenterX;
			sy = Math.Sin((uSec * curSec) - HalfPi) * 50 + (double)CenterY;

			// Ici c'est le calcul de l'origine de la trotteuse qui n'est pas au
			// centre de l'image mais 30 pixels plus loin (on ne retire pas pi/2 
			// mais on le rajoute car il est bien de l'autre coté de l'origine du repere
			osx = Math.Cos((uSec * curSec) + HalfPi) * 20 + (double)CenterX;
			osy = Math.Sin((uSec * curSec) + HalfPi) * 20 + (double)CenterY;

			//calcul des minutes
			mmx = Math.Cos((uSec * curMin) - HalfPi) * 40 + (double)CenterX;
			mmy = Math.Sin((uSec * curMin) - HalfPi) * 40 + (double)CenterY;
			omx = Math.Cos((uSec * curMin) + HalfPi) * 20 + (double)CenterX;
			omy = Math.Sin((uSec * curMin) + HalfPi) * 20 + (double)CenterY;

			// idem mais avec l'angle representant 1 heure	
			hx = Math.Cos((uHour * curHour) - HalfPi) * 45 + (double)CenterX;
			hy = Math.Sin((uHour * curHour) - HalfPi) * 45 + (double)CenterY;
			ohx = Math.Cos((uHour * curHour) + HalfPi) * 20 + (double)CenterX;
			ohy = Math.Sin((uHour * curHour) + HalfPi) * 20 + (double)CenterY;

			// le padleft sert à forcer l'affichage de la valeur sur 2 digits
			// par exple : 11:1:8 devient 11:01:08 ce qui est + esthétique :)
			label1.Text = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" +
						  DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" +
						  DateTime.Now.Second.ToString().PadLeft(2, '0');

			/* C'est ici qu'on dessine
			 * on commence par l'image de fond qui va tout ecraser (le clear fait scintiller c'est génant)
			 * ensuite dans l'ordre dans lequel on veut que les aiguilles se surperposent.
			 * En général, on commence par la trotteuse, ensuite celle des minutes et enfin pour les heures
			 */
			// j'ai effectivement mis l'image de fond en background de ma forme
			// on pourrait tout aussi bien la mettre dans une imagelist
			formGraphics.DrawImage(this.BackgroundImage, pictureBox1.Bounds);

			// Définition de la couleur d'ecriture, largeur
			myPen.Color = System.Drawing.Color.White;
			myPen.Width = 3;
			// minute
			formGraphics.DrawLine(myPen, (int)omx, (int)omy, (int)mmx, (int)mmy);
			// heure
			formGraphics.DrawLine(myPen, (int)ohx, (int)ohy, (int)hx, (int)hy);

			// Petit bonus, j'ai dessiné un petit cercle pour le centre
			myPen.Width = 2;
			myPen.Color = System.Drawing.Color.White;
			formGraphics.FillEllipse(myBrush, CenterX - 4, CenterY - 4, 8, 8);

			// c'est pour le contour ^_^
			formGraphics.DrawEllipse(myPen, CenterX - 4, CenterY - 4, 8, 8);

			// la trotteuse
			// Définition de la couleur d'ecriture, largeur
			myPen.Color = System.Drawing.Color.LightSkyBlue;
			myPen.Width = 2;//2
			formGraphics.DrawLine(myPen, (int)CenterX, (int)CenterY, (int)sx, (int)sy);
			formGraphics.DrawLine(myPen, (int)osx, (int)osy, (int)CenterX, (int)CenterY);
		}
	}
}
