using Novacode;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GenerateCV
{
    public partial class FeedFields : Form
    {
        public FeedFields()
        {
            InitializeComponent();
            InitializeKeyWord();
        }

        List<string> keyWord = new List<string>();
        void InitializeKeyWord()
        {
            keyWord.Add("NomPrenom");
            keyWord.Add("Grade");
            keyWord.Add("Pole");
            keyWord.Add("Titre");
            keyWord.Add("Specialite");
            keyWord.Add("Langue");
            keyWord.Add("CompTech1");
            keyWord.Add("CompTech2");
            keyWord.Add("CompTech3");
            keyWord.Add("CompTech4");
            keyWord.Add("CompTech5");
            keyWord.Add("CompTech6");
            keyWord.Add("CompTech7");
            keyWord.Add("CompTech8");
            keyWord.Add("CompTech9");
            keyWord.Add("CompTechs1");
            keyWord.Add("CompTechs2");
            keyWord.Add("Metier1");
            keyWord.Add("Metier2");
            keyWord.Add("Metier3");
            keyWord.Add("Metier4");
            keyWord.Add("Metier5");
            keyWord.Add("Metier6");
            keyWord.Add("Formations1");
            keyWord.Add("Formations2");
            keyWord.Add("Formations3");
            keyWord.Add("Formations4");
            keyWord.Add("DetailFormation1");
            keyWord.Add("DetailFormation2");
            keyWord.Add("DetailFormation3");
            keyWord.Add("DetailFormation4");
            keyWord.Add("FormComp1");
            keyWord.Add("FormComp2");
            keyWord.Add("FormComp3");
            keyWord.Add("FormComp4");
            keyWord.Add("Certif1");
            keyWord.Add("Certif2");
            keyWord.Add("Certif3");
            keyWord.Add("Certif4");
            keyWord.Add("TitreMissionUn");
            keyWord.Add("DureeMisionUn");
            keyWord.Add("DetailMissionUn");
            keyWord.Add("EnvTechniqueUn");
            keyWord.Add("TitreMissionDeux");
            keyWord.Add("DureeMisionDeux");
            keyWord.Add("DetailMissionDeux");
            keyWord.Add("EnvTechniqueDeux");
            keyWord.Add("TitreMissionTrois");
            keyWord.Add("DureeMisionTrois");
            keyWord.Add("DetailMissionTrois");
            keyWord.Add("EnvTechniqueTrois");
            keyWord.Add("TitreMissionQuatre");
            keyWord.Add("DureeMissionQuatre");
            keyWord.Add("DetailMissionQuatre");
            keyWord.Add("EnvTechniqueQuatre");          
            keyWord.Add("TitreMissionCinq");
            keyWord.Add("DureeMisionCinq");
            keyWord.Add("DetailMissionCinq");
            keyWord.Add("EnvTechniqueCinq");
            keyWord.Add("Environnement technique Un:");
            keyWord.Add("Environnement technique Deux:");
            keyWord.Add("Environnement technique Trois:");
            keyWord.Add("Environnement technique Quatre:");
            keyWord.Add("Environnement technique Cinq:");
            keyWord.Add("Environnement technique Six:");
            keyWord.Add("Environnement fonctionnel Un:");
            keyWord.Add("Environnement fonctionnel Deux:");
            keyWord.Add("Environnement fonctionnel Trois:");
            keyWord.Add("Environnement fonctionnel Quatre:");
            keyWord.Add("Environnement fonctionnel Cinq:");
            keyWord.Add("Environnement fonctionnel Six:");
            keyWord.Add("EnvFonctUn");
            keyWord.Add("EnvFonctDeux");
            keyWord.Add("EnvFonctTrois");
            keyWord.Add("EnvFonctQuatre");
            keyWord.Add("EnvFonctCinq");
            keyWord.Add("EnvFonctSix");
            keyWord.Add("[Compétences techniques]");
            keyWord.Add("[Compétences finances]");
            keyWord.Add("[Formation]");
            keyWord.Add("[Certification]");
            keyWord.Add("[FormationComplementaire]");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(outPut.Text))
            {
                MessageBox.Show("Veuillez sélectionner un répertoire de sortie", "Alert");
                outPut.BackColor = Color.Red;
                return;
            }            
         

            DocX template = DocX.Load(Path.Combine(Directory.GetCurrentDirectory(), "TemplateCV.docx"));

            PopulateInfos(template);
            PopulateTechSkills(template);
            PopulateMetierSkills(template);
            PopulateFormation(template);
            PopulateExperience(template);

            Clear(template);
            var fileName = System.IO.Path.Combine(outPut.Text, string.Format("{0}{1}",string.IsNullOrEmpty(name.Text)? "CV": name.Text.Trim().Replace(" ",""), ".docx"));
            try
            {
                template.SaveAs(fileName);
                MessageBox.Show(string.Format("Génération terminée dans : {0}", fileName),"Info");
                Process.Start("WINWORD.EXE", fileName);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }           
        }

        private void PopulateExperience(DocX template)
        {
            template.ReplaceText("TitreMissionUn", titreM1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionUn", detailM1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionUn", dureeM1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueUn", envT1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctUn", envF1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 1);
            if (!string.IsNullOrEmpty(envT1.Text))
                template.ReplaceText("Environnement technique Un:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF1.Text))
                template.ReplaceText("Environnement fonctionnel Un:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);


            template.ReplaceText("TitreMissionDeux", titreM2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionDeux", detailM2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionDeux", dureeM2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctDeux", envF2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueDeux", envT2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 2);
            if (!string.IsNullOrEmpty(envT2.Text))
                template.ReplaceText("Environnement technique Deux:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF2.Text))
                template.ReplaceText("Environnement fonctionnel Deux:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);


            template.ReplaceText("TitreMissionTrois", titreM3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionTrois", detailM3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionTrois", dureeM3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueTrois", envT3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctTrois", envF3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 3);
            if (!string.IsNullOrEmpty(envT3.Text))
                template.ReplaceText("Environnement technique Trois:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF3.Text))
                template.ReplaceText("Environnement fonctionnel Trois:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            template.ReplaceText("TitreMissionQuatre", titreM4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionQuatre", detailM4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionQuatre", dureeM4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueQuatre", envT4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctQuatre", envF4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 4);
            if (!string.IsNullOrEmpty(envT4.Text))
                template.ReplaceText("Environnement technique Quatre:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF4.Text))
                template.ReplaceText("Environnement fonctionnel Quatre:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            template.ReplaceText("TitreMissionCinq", titreM5.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionCinq", detailM5.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionCinq", dureeM5.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueCinq", envT5.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctCinq", envF5.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 5);
            if (!string.IsNullOrEmpty(envT5.Text))
                template.ReplaceText("Environnement technique Cinq:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF5.Text))
                template.ReplaceText("Environnement fonctionnel Cinq:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            template.ReplaceText("TitreMissionSix", titreM6.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailMissionSix", detailM6.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DureeMissionSix", dureeM6.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvTechniqueSix", envT6.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("EnvFonctSix", envF6.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            PopulateTask(template, 6);
            if (!string.IsNullOrEmpty(envT6.Text))
                template.ReplaceText("Environnement technique Six:", "Environnement technique : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            if (!string.IsNullOrEmpty(envF6.Text))
                template.ReplaceText("Environnement fonctionnel Six:", "Environnement fonctionnel : ", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
        }

        private void PopulateFormation(DocX template)
        {
            template.ReplaceText("Formations1", formation1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Formations2", formation2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Formations3", formation3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Formations4", formation4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailFormation1", detailForm1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailFormation2", detailForm2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailFormation3", detailForm3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("DetailFormation4", detailForm4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Certif1", certif1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Certif2", certif2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Certif3", certif3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Certif4", certif4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            
            template.ReplaceText("FormComp1", formComp1.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("FormComp2", formComp2.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("FormComp3", formComp3.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("FormComp4", formComp4.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            int i = 1;
            while (i <= 4)
            {
                RichTextBox tbx = this.Controls.Find(string.Format("formation{0}", i), true).FirstOrDefault() as RichTextBox;
                if (!string.IsNullOrEmpty(tbx.Text))
                {
                    template.ReplaceText("[Formation]", "Formations", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                    break;
                }
                i++;
            }

            i = 1;
            while (i <= 4)
            {
                RichTextBox tbx = this.Controls.Find(string.Format("certif{0}", i), true).FirstOrDefault() as RichTextBox;
                if (!string.IsNullOrEmpty(tbx.Text))
                {
                    template.ReplaceText("[Certification]", "Certifications", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                    break;
                }
                i++;
            }

            i = 1;
            while (i <= 4)
            {
                RichTextBox tbx = this.Controls.Find(string.Format("formComp{0}", i), true).FirstOrDefault() as RichTextBox;
                if (!string.IsNullOrEmpty(tbx.Text))
                {
                    template.ReplaceText("[FormationComplementaire]", "Formations Complementaires", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                    break;
                }
                i++;
            }           
        }

        private void PopulateMetierSkills(DocX template)
        {
            template.ReplaceText("Metier1", metier1.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier2", metier2.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier3", metier3.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier4", metier4.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier5", metier5.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier6", metier6.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Metier7", metier7.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            int i = 1;
            while (i <= 7)
            {
                TextBox tbx = this.Controls.Find(string.Format("metier{0}", i), true).FirstOrDefault() as TextBox;
                if (!string.IsNullOrEmpty(tbx.Text))
                {
                    template.ReplaceText("[Compétences finances]", "Compétences fonctionelles", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                    break;
                }
                i++;
            }
        }

        private void PopulateTechSkills(DocX template)
        {
            template.ReplaceText("CompTech1", technique1.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech2", technique2.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech3", technique3.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech4", technique4.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech5", technique5.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech6", technique6.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech7", technique7.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech8", technique8.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTech9", technique9.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTechs1", technique10.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("CompTechs2", technique11.Text, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            int i = 1;
            while(i<=11)           
            {
                TextBox tbx = this.Controls.Find(string.Format("technique{0}", i), true).FirstOrDefault() as TextBox;
                if (!string.IsNullOrEmpty(tbx.Text))
                {
                    template.ReplaceText("[Compétences techniques]", "Compétences techniques", false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                    break;
                }
                i++;
            }
            
    }

        private void PopulateInfos(DocX template)
        {
            template.ReplaceText("Profile", poste.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Specialite", specialite.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("NomPrenom", name.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Grade", grade.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Pole", pole.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("[Presentation]", presentation.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            template.ReplaceText("Langue", langue.Text.TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);            
    }

        private void PopulateTask(DocX template, int missionId)
        {
            RichTextBox tbx = this.Controls.Find(string.Format("task{0}", missionId), true).FirstOrDefault() as RichTextBox;
            if (!tbx.Lines.Any()) return;
            var i = 0;
            var indice = tbx.Lines.Length < 10 ? tbx.Lines.Length : 9;
            while (i < indice)
            {
                template.ReplaceText(string.Format("[Mission{0}Task{1}]", missionId, i), tbx.Lines[i].TrimEnd().Value(), false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
                i++;
            }           
        }

        private void Clear(DocX template)
        {
            foreach (var key in keyWord)
                template.ReplaceText(key, string.Empty, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);

            for(int i=1; i<7;i++)
            {
                for(int j=0; j<10;j++)
                    template.ReplaceText(string.Format("[Mission{0}Task{1}]", i, j), string.Empty, false, System.Text.RegularExpressions.RegexOptions.CultureInvariant, null, null, MatchFormattingOptions.ExactMatch);
            }
        }

        private void exp1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void sep1_Click(object sender, EventArgs e)
        {

        }

        private void exp2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }
        

        private void AddExp1_Click_1(object sender, EventArgs e)
        {            
            AddExp1.Visible = false;
            sep1.Visible = exp2.Visible = true;
            tabPage5.AutoScrollPosition = exp2.Location;
        }

        private void AddExp2_Click(object sender, EventArgs e)
        {
            AddExp2.Visible = false;
            sep2.Visible = exp3.Visible = true;
            tabPage5.AutoScrollPosition = exp3.Location;
        }

        private void AddExp3_Click(object sender, EventArgs e)
        {
            AddExp3.Visible = false;
            sep3.Visible = exp4.Visible = true;
            tabPage5.AutoScrollPosition = exp4.Location;
        }

        private void AddExp4_Click_1(object sender, EventArgs e)
        {
            AddExp4.Visible = false;
            sep4.Visible = exp5.Visible = true;
            tabPage5.AutoScrollPosition = exp5.Location;
        }

        private void AddExp5_Click_1(object sender, EventArgs e)
        {
            AddExp5.Visible = false;
            sep5.Visible = exp6.Visible = true;
            tabPage5.AutoScrollPosition = exp6.Location;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            DialogResult result = fbd.ShowDialog();

            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                outPut.Text = fbd.SelectedPath;
                outPut.BackColor = Color.Wheat;
            }
        }
    }
}
