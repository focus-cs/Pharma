/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.sciforma.pharma;

import com.sciforma.pharma.beans.Connector;
import com.sciforma.pharma.beans.SciformaField;
import com.sciforma.pharma.manager.ProjectManager;
import com.sciforma.pharma.manager.ProjectManagerImpl;
import com.sciforma.pharma.util.CSVUtils;
import com.sciforma.psnext.api.CostAssignment;
import com.sciforma.psnext.api.DatedData;
import com.sciforma.psnext.api.DoubleDatedData;
import com.sciforma.psnext.api.PSException;
import com.sciforma.psnext.api.Project;
import com.sciforma.psnext.api.ResAssignment;
import com.sciforma.psnext.api.Session;
import com.sciforma.psnext.api.Task;
import com.sciforma.psnext.api.TaskOutlineList;
import com.sciforma.psnext.api.Transaction;
import java.io.FileWriter;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.pmw.tinylog.Logger;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.FileSystemXmlApplicationContext;

/**
 *
 * @author lahou
 */
public class APIPharma {

    private static final String VERSION = "1.13";

    public static ApplicationContext ctx;

    private static String IP;
    private static String PORT;
    private static String CONTEXTE;
    private static String USER;
    private static String PWD;

    public static final char DEFAULT_SEPARATOR = ';';
    private static String CSV_FILE;

    public static Session mSession;
    private static ProjectManager projectManager;

    private static final long CONST_DURATION_OF_DAY = 1000 * 60 * 60 * 24;

    public static void main(String[] args) {
        Logger.info("[main][V" + VERSION + "] Demarrage de l'API: " + new Date());
        try {
            initialisation();
            connexion();
            chargementParametreSciforma();
            process();
            mSession.logout();
            Logger.info("[main][V" + VERSION + "] Fin de l'API: " + new Date());
        } catch (PSException ex) {
            Logger.error(ex);
        }
        System.exit(0);
    }

    private static void initialisation() {
        ctx = new FileSystemXmlApplicationContext(System.getProperty("user.dir") + System.getProperty("file.separator") + "config" + System.getProperty("file.separator") + "applicationContext.xml");
    }

    private static void connexion() {
        Logger.info("[connexion] Chargement parametre connexion:" + new Date());
        Connector c = (Connector) ctx.getBean("sciforma");
        USER = c.getUSER();
        PWD = c.getPWD();
        IP = c.getIP();
        PORT = c.getPORT();
        CONTEXTE = c.getCONTEXTE();

        try {
            Logger.info("[connexion] Initialisation de la Session:" + new Date());
            String url = IP + "/" + CONTEXTE;
            Logger.info("[connexion] " + url + " :" + new Date());
            mSession = new Session(url);
            mSession.login(USER, PWD.toCharArray());
            Logger.info("[connexion] Connecté: " + new Date() + " Ă  l'instance " + CONTEXTE);
        } catch (PSException ex) {
            Logger.error("[connexion] Erreur dans la connection de ... " + CONTEXTE, ex);
            System.exit(-1);
        }
    }

    private static void chargementParametreSciforma() throws PSException {
        try {
            Logger.info("Demarrage du chargement des parametres de l'application:" + new Date());
            projectManager = new ProjectManagerImpl(mSession);
            CSV_FILE = ((SciformaField) ctx.getBean("sciforma_to_csv")).getSciformaField();
            //FILENAME_IMPORT_SAP = ((SciformaField) ctx.getBean("sap_to_sciforma")).getSciformaField();
            Logger.info("Fin du chargement des parametres de l'application:" + new Date());
        } catch (Exception ex) {
            Logger.error("Erreur dans la lecture l'intitialisation du parametrage " + new Date(), ex);
            mSession.logout();
            System.exit(1);
        }
    }

    private static void process() throws PSException {
        try {
            FileWriter writer = new FileWriter(CSV_FILE);
            NumberFormat nf = new DecimalFormat("######.####");
            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");
            List<String> valueToWrite = new ArrayList<>();
            CSVUtils.writeLine(writer, Arrays.asList("Periode Start", "Periode Finish", "Year", "Quarter", "Month", "Week",
                    "Line Type", "Object Name", "Transaction Nature", "Transaction Type",
                    "Transaction status", "Organisation", "Object ID", "Resource Manager1",
                    "Remaining Effort/Cost", "Actual Effort/Cost", "Total Effort/Cost", "Amount",
                    "Hrcosts", "Job Classification", "#", "Tag Zone", "Tag project Scope",
                    "Tag Portfolio Scope", "Tag departement", "Tag coutries", "Tag major contry in zone",
                    "Tag chapter", "Study Code", "TaskID", "TaskName", "Type", "Hierarchy N-1",
                    "Hierarchy N-2", "Hierarchy N-3", "Hierarchy N-4", "Hierarchy N-5", "Hierarchy N-6",
                    "Hierarchy N-7", "Project Id", "Project Name", "Priority", "STG", "Manager 1",
                    "Product Code", "Project Code", "Project Type", "Category", "Franchise",
                    "Target species", "Therapeutic axis", "Intendedindication", "Innovation level",
                    "Targeted zones", "Strategic Product", "API_activeingredie", "Dosage form",
                    "Portfolio Folder", "Project owner", "Last Modification", "Last published Date",
                    "Coordinator", "Workflow State", "Project Frizez phase", "Project phase",
                    "OfferConceptionman", "Patent team member", "POC process pilot",
                    "Galenicprocesspilo", "Industrializationp", "Safety process pilot", "Dose_Indicationpro",
                    "RegAffairsteamlead", "Marketingteammembe", "SupplyChainteammem", "PerformingOrganiza",
                    "Owning Organization", "Users readers", "Users writers", "Organization readers",
                    "Organization writers", "Core team readers", "Core team writers",
                    "Targeted countries", "Workpackage_ID", "Workpackage_Manager 1",
                    "Workpackage_Manager 2", "Workpackage_Manager 3", "Workpackage_Name",
                    "Workpackage_Priority", "Workpackage_Userwriter",
                    "TO Year 5 Countries", "TO Year 5 k€", "TO Year 5 Market value", "TO Year 5 Zones"));

            List<Project> lp = mSession.getProjectList(Project.VERSION_WORKING, Project.READWRITE_ACCESS);
            int nbProjet = lp.size();
            Iterator lpit = lp.iterator();
            while (lpit.hasNext()) {
                Project p = (Project) lpit.next();
                Logger.info("=======================================================================================");
                Logger.info("Traitement du projet [" + (lp.indexOf(p) + 1) + "/" + nbProjet + "] " + p.getStringField("Name"));
                Logger.info("=======================================================================================");
                try {
                    p.open(true);
                    if (p.getBooleanField("CEVA - API Export Filter")) {
                        Logger.info("                    " + p.getStringField("Name") + " => Réponds au critère du filtre CEVA - API Export Filter");
                        TaskOutlineList tasks = p.getTaskOutlineList();
                        Iterator taskIt = tasks.iterator();
                        while (taskIt.hasNext()) {
                            Task task = (Task) taskIt.next();

                            Calendar cDebut = Calendar.getInstance();
                            Calendar cDebutFin = Calendar.getInstance();
                            Calendar cFin = Calendar.getInstance();

                            cDebut.setTime(task.getDateField("Start"));
                            cDebutFin.setTime(task.getDateField("Start"));
                            cDebutFin.add(Calendar.DATE, 6);
                            cFin.setTime(task.getDateField("Finish"));

                            Logger.info("Task:" + task.getStringField("name"));
                            if (task.getBooleanField("CEVA - API Export Filter")) {
                                if (task.getResAssignmentList().size() > 0) {
                                    //**************************************************************************************\\
                                    Logger.info("   Resource assignment(s) to task '" + task.getStringField("name") + "':");
                                    //**************************************************************************************\\
                                    Iterator resAssignIt = task.getResAssignmentList().iterator();
                                    while (resAssignIt.hasNext()) {
                                        cDebut.setTime(task.getDateField("Start"));
                                        cDebutFin.setTime(task.getDateField("Start"));
                                        cDebutFin.add(Calendar.DATE, 6);
                                        cFin.setTime(task.getDateField("Finish"));
                                        ResAssignment resAssign = (ResAssignment) resAssignIt.next();
                                        if (resAssign.getBooleanField("CEVA - API Export Filter")) {
                                            Logger.info("      " + resAssign.getStringField("name"));
                                            while (cDebut.before(cFin)) {
                                                valueToWrite.add(sdf.format(cDebut.getTime()));
                                                valueToWrite.add(sdf.format(cDebutFin.getTime()));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.YEAR)));
                                                valueToWrite.add((cDebut.get(Calendar.MONTH) >= Calendar.JANUARY && cDebut.get(Calendar.MONTH) <= Calendar.MARCH) ? "Q1" : (cDebut.get(Calendar.MONTH) >= Calendar.APRIL && cDebut.get(Calendar.MONTH) <= Calendar.JUNE) ? "Q2" : (cDebut.get(Calendar.MONTH) >= Calendar.JULY && cDebut.get(Calendar.MONTH) <= Calendar.SEPTEMBER) ? "Q3" : "Q4");
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.MONTH) + 1));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.WEEK_OF_YEAR) == 53 ? 1 : cDebut.get(Calendar.WEEK_OF_YEAR) + 1));
                                                valueToWrite.add("Resource");
                                                valueToWrite.add(resAssign.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(""); //TRANSACTION Nature…
                                                valueToWrite.add(""); //transaction type
                                                valueToWrite.add(""); //transaction status
                                                valueToWrite.add(resAssign.getStringField("Resource organization").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(resAssign.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(resAssign.getStringField("API Manager 1").replace("\n", ",").replace(";", ",")); //Resourcemanager1 ? formule rajouté

                                                List<DoubleDatedData> lddd = resAssign.getDatedData("Remaining Effort", DatedData.DAY, cDebut.getTime(), cDebutFin.getTime());
                                                double data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                lddd = resAssign.getDatedData("Actual Effort", DatedData.DAY, cDebut.getTime(), cDebutFin.getTime());
                                                data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                lddd = resAssign.getDatedData("Total Effort", DatedData.DAY, cDebut.getTime(), cDebutFin.getTime());
                                                data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                valueToWrite.add(""); // Amount
                                                /*lddd = resAssign.getDatedData("Hrcosts_k€/d", DatedData.DAY, cDebut.getTime(), cDebutFin.getTime());
                                                data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }*/
                                                valueToWrite.add(nf.format(resAssign.getDoubleField("Hrcosts_k€/d")*8).replace(".", ","));

                                                valueToWrite.add(resAssign.getStringField("Job Classification").replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(String.valueOf(task.getIntField("#")).replace("\n", ",").replace(";", ","));

                                                List<String> ls = task.getListField("Tag Zone");
                                                String text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Project Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getStringField("Tag PMO Scope").replace("\n", ",").replace(";", ",")); //Tag Portfolio Scope?

                                                ls = task.getListField("Tag Department Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getBooleanField("Tag Major Country in Zone") ? "true" : "false");
                                                valueToWrite.add(task.getStringField("Tag Chapter").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("PHARMA - Study Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getBooleanField("milestones") ? "Milestones" : "Task");
                                                valueToWrite.add(task.getStringField("Hierarchy N-1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-2").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-3").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-4").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-5").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-6").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-7").replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("priority").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Strategic Target Group");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Product code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Type").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Type").replace("\n", ",").replace(";", ",")); //Project Type / Category identique ?
                                                valueToWrite.add(p.getStringField("Franchise").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Target species");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Therapeutic axis");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Intended indication(s)").replace("\n", ",").replace(";", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Innovation level").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Targeted zones");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getBooleanField("Strategic Product") ? "true" : "false");

                                                ls = p.getListField("API (active ingredient)");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Dosage form").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Portfolio Folder").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project owner").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(sdf.format(p.getDateField("Modified Date")).replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(sdf.format(p.getDateField("Published Date")).replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Coordinator").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Workflow State").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Frise Phase project").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project phase").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Offer Conception manager").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Patent team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("POC process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Galenic process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Industrialization process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Safety process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Dose-Indication process pilot").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Reg Affairs team leaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Marketing team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Supply Chain team member").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Performing Organizations");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Owning Organization").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("usersReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("usersWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Readers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Writers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Targeted countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                if (!task.getStringField("Work Package ID").isEmpty()) {
                                                    valueToWrite.add(task.getStringField("Work Package ID").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 2").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 3").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("workPackagePriority").replace("\n", ",").replace(";", ","));
                                                    ls = task.getListField("workPackageUserWriter");
                                                    text = "";
                                                    for (String l : ls) {
                                                        if (!text.equals("")) {
                                                            text += ",";
                                                        }
                                                        text += l;
                                                    }
                                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                } else {
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                }

                                                ls = task.getListField("TO Year 5 Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 k€")).replace(".", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 Market value")).replace(".", ","));
                                                valueToWrite.add(task.getStringField("TO Year 5 Zones").replace("\n", ",").replace(";", ","));

                                                CSVUtils.writeLine(writer, valueToWrite);
                                                valueToWrite = new ArrayList<>();
                                                cDebut.add(Calendar.DATE, 7);
                                                cDebutFin.add(Calendar.DATE, 7);
                                            }
                                        }
                                    }
                                }

                                cDebut.setTime(task.getDateField("Start"));
                                cDebutFin.setTime(task.getDateField("Start"));
                                cDebutFin.add(Calendar.DATE, 6);
                                cFin.setTime(task.getDateField("Finish"));
                                if (task.getCostAssignmentList().size() > 0) {
                                    //**************************************************************************************\\
                                    Logger.info("   Cost assignment(s) to task '" + task.getStringField("name") + "':");
                                    //**************************************************************************************\\
                                    Iterator costAssignIt = task.getCostAssignmentList().iterator();
                                    while (costAssignIt.hasNext()) {
                                        cDebut.setTime(task.getDateField("Start"));
                                        cDebutFin.setTime(task.getDateField("Start"));
                                        cDebutFin.add(Calendar.DATE, 6);
                                        cFin.setTime(task.getDateField("Finish"));
                                        CostAssignment costAssign = (CostAssignment) costAssignIt.next();
                                        if (costAssign.getBooleanField("CEVA - API Export Filter")) {
                                            Logger.info("      " + costAssign.getStringField("name"));
                                            while (cDebut.before(cFin)) {
                                                valueToWrite.add(sdf.format(cDebut.getTime()));
                                                valueToWrite.add(sdf.format(cDebutFin.getTime()));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.YEAR)));
                                                valueToWrite.add((cDebut.get(Calendar.MONTH) >= Calendar.JANUARY && cDebut.get(Calendar.MONTH) <= Calendar.MARCH) ? "Q1" : (cDebut.get(Calendar.MONTH) >= Calendar.APRIL && cDebut.get(Calendar.MONTH) <= Calendar.JUNE) ? "Q2" : (cDebut.get(Calendar.MONTH) >= Calendar.JULY && cDebut.get(Calendar.MONTH) <= Calendar.SEPTEMBER) ? "Q3" : "Q4");
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.MONTH) + 1));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.WEEK_OF_YEAR) == 53 ? 1 : cDebut.get(Calendar.WEEK_OF_YEAR) + 1));
                                                valueToWrite.add("Cost Item");
                                                valueToWrite.add(costAssign.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(costAssign.getStringField("Effective Organization").replace("\n", ",").replace(";", ",")); //TRANSACTION Nature…?
                                                valueToWrite.add(costAssign.getStringField("Organization").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(costAssign.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add("");
                                                valueToWrite.add(""); //transaction type
                                                valueToWrite.add(""); //transaction type
                                                List<DoubleDatedData> lddd = costAssign.getDatedData("Remaining Cost", DatedData.DAY, cDebut.getTime(), cDebutFin.getTime());
                                                double data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                lddd = costAssign.getDatedData("Actual Cost", DatedData.NONE, cDebut.getTime(), cDebutFin.getTime());
                                                data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                lddd = costAssign.getDatedData("Total Cost", DatedData.NONE, cDebut.getTime(), cDebutFin.getTime());
                                                data = 0;
                                                for (DoubleDatedData ddd : lddd) {
                                                    data = data + ddd.getData();
                                                }
                                                valueToWrite.add(nf.format(data).replace(".", ","));

                                                valueToWrite.add(""); // Amount
                                                valueToWrite.add(""); // Hrcosts_k€/d
                                                valueToWrite.add(costAssign.getStringField("Category"));

                                                valueToWrite.add(String.valueOf(task.getIntField("#")));

                                                List<String> ls = task.getListField("Tag Zone");
                                                String text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Project Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getStringField("Tag PMO Scope").replace("\n", ",").replace(";", ",")); //Tag Portfolio Scope?

                                                ls = task.getListField("Tag Department Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getBooleanField("Tag Major Country in Zone") ? "true" : "false");
                                                valueToWrite.add(task.getStringField("Tag Chapter").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("PHARMA - Study Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getBooleanField("milestones") ? "Milestones" : "Task");
                                                valueToWrite.add(task.getStringField("Hierarchy N-1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-2").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-3").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-4").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-5").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-6").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-7").replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("priority").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Strategic Target Group");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Product code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Type").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Type").replace("\n", ",").replace(";", ",")); //Project Type / Category identique ?
                                                valueToWrite.add(p.getStringField("Franchise").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Target species");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Therapeutic axis");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Intended indication(s)").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Innovation level").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Targeted zones");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getBooleanField("Strategic Product") ? "true" : "false");

                                                ls = p.getListField("API (active ingredient)");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Dosage form").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Portfolio Folder").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project owner").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(sdf.format(p.getDateField("Modified Date")));
                                                valueToWrite.add(sdf.format(p.getDateField("Published Date")));
                                                valueToWrite.add(p.getStringField("Coordinator").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Workflow State").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Frise Phase project").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project phase").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Offer Conception manager").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Patent team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("POC process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Galenic process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Industrialization process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Safety process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Dose-Indication process pilot").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Reg Affairs team leaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Marketing team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Supply Chain team member").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Performing Organizations");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Owning Organization").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("usersReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("usersWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Readers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Writers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Targeted countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                if (!task.getStringField("Work Package ID").isEmpty()) {
                                                    valueToWrite.add(task.getStringField("Work Package ID").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 2").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 3").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("workPackagePriority").replace("\n", ",").replace(";", ","));
                                                    ls = task.getListField("workPackageUserWriter");
                                                    text = "";
                                                    for (String l : ls) {
                                                        if (!text.equals("")) {
                                                            text += ",";
                                                        }
                                                        text += l;
                                                    }
                                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                } else {
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                }

                                                ls = task.getListField("TO Year 5 Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 k€")).replace(".", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 Market value")).replace(".", ","));
                                                valueToWrite.add(task.getStringField("TO Year 5 Zones").replace("\n", ",").replace(";", ","));

                                                CSVUtils.writeLine(writer, valueToWrite);
                                                valueToWrite = new ArrayList<>();
                                                cDebut.add(Calendar.DATE, 7);
                                                cDebutFin.add(Calendar.DATE, 7);
                                            }
                                        }
                                    }
                                }

                                cDebut.setTime(task.getDateField("Start"));
                                cDebutFin.setTime(task.getDateField("Start"));
                                cDebutFin.add(Calendar.DATE, 6);
                                cFin.setTime(task.getDateField("Finish"));
                                long diff = Math.abs(cDebut.getTime().getTime() - cFin.getTime().getTime());
                                long numberOfDay = (long) diff / CONST_DURATION_OF_DAY;
                                int numberWeek = (int) Math.ceil(numberOfDay / 7);
                                numberWeek++;
                                if (task.getTransactionList().size() > 0) {
                                    cDebut.setTime(task.getDateField("Start"));
                                    cDebutFin.setTime(task.getDateField("Start"));
                                    cDebutFin.add(Calendar.DATE, 6);
                                    cFin.setTime(task.getDateField("Finish"));
                                    //**************************************************************************************\\
                                    Logger.info("   Transaction(s) to task '" + task.getStringField("name") + "':");
                                    //**************************************************************************************\\
                                    Iterator transactionIt = task.getTransactionList().iterator();
                                    while (transactionIt.hasNext()) {
                                        Transaction transaction = (Transaction) transactionIt.next();
                                        if (transaction.getBooleanField("CEVA - API Export Filter")) {
                                            Logger.info("      " + transaction.getStringField("name"));
                                            while (cDebut.before(cFin)) {
                                                valueToWrite.add(sdf.format(cDebut.getTime()));
                                                valueToWrite.add(sdf.format(cDebutFin.getTime()));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.YEAR)));
                                                valueToWrite.add((cDebut.get(Calendar.MONTH) >= Calendar.JANUARY && cDebut.get(Calendar.MONTH) <= Calendar.MARCH) ? "Q1" : (cDebut.get(Calendar.MONTH) >= Calendar.APRIL && cDebut.get(Calendar.MONTH) <= Calendar.JUNE) ? "Q2" : (cDebut.get(Calendar.MONTH) >= Calendar.JULY && cDebut.get(Calendar.MONTH) <= Calendar.SEPTEMBER) ? "Q3" : "Q4");
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.MONTH) + 1));
                                                valueToWrite.add(String.valueOf(cDebut.get(Calendar.WEEK_OF_YEAR) == 53 ? 1 : cDebut.get(Calendar.WEEK_OF_YEAR) + 1));
                                                valueToWrite.add("Transaction");
                                                valueToWrite.add(transaction.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(transaction.getStringField("nature").replace("\n", ",").replace(";", ",")); //TRANSACTION Nature…?
                                                valueToWrite.add(transaction.getStringField("Type").replace("\n", ",").replace(";", ",")); //transaction type
                                                valueToWrite.add(transaction.getStringField("Status").replace("\n", ",").replace(";", ",")); //transaction status
                                                valueToWrite.add("");//valueToWrite.add(transaction.getStringField("Organization"));
                                                valueToWrite.add(transaction.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add("");
                                                valueToWrite.add("");
                                                valueToWrite.add("");
                                                valueToWrite.add("");
                                                double amount = transaction.getDoubleField("Amount  API");
                                                double amountPerWeek = amount / numberWeek;
                                                valueToWrite.add(nf.format(amountPerWeek).replace(".", ","));
                                                valueToWrite.add(""); // Hrcosts_k€/d
                                                valueToWrite.add("");
                                                valueToWrite.add(String.valueOf(task.getIntField("#")));

                                                List<String> ls = task.getListField("Tag Zone");
                                                String text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Project Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getStringField("Tag PMO Scope").replace("\n", ",").replace(";", ",")); //Tag Portfolio Scope?

                                                ls = task.getListField("Tag Department Scope");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = task.getListField("Tag Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(task.getBooleanField("Tag Major Country in Zone") ? "true" : "false");
                                                valueToWrite.add(task.getStringField("Tag Chapter").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("PHARMA - Study Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getBooleanField("milestones") ? "Milestones" : "Task");
                                                valueToWrite.add(task.getStringField("Hierarchy N-1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-2").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-3").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-4").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-5").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-6").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(task.getStringField("Hierarchy N-7").replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("ID").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("priority").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Strategic Target Group");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Product code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Code").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project Type").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Type").replace("\n", ",").replace(";", ",")); //Project Type / Category identique ?
                                                valueToWrite.add(p.getStringField("Franchise").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Target species");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Therapeutic axis");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Intended indication(s)").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Innovation level").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Targeted zones");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getBooleanField("Strategic Product") ? "true" : "false");

                                                ls = p.getListField("API (active ingredient)");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Dosage form").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Portfolio Folder").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project owner").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(sdf.format(p.getDateField("Modified Date")));
                                                valueToWrite.add(sdf.format(p.getDateField("Published Date")));
                                                valueToWrite.add(p.getStringField("Coordinator").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Workflow State").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Frise Phase project").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Project phase").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Offer Conception manager").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Patent team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("POC process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Galenic process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Industrialization process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Safety process pilot").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Dose-Indication process pilot").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Reg Affairs team leaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                                valueToWrite.add(p.getStringField("Marketing team member").replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Supply Chain team member").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("Performing Organizations");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(p.getStringField("Owning Organization").replace("\n", ",").replace(";", ","));

                                                ls = p.getListField("usersReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("usersWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationReaders");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("organizationWriters");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Readers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Core Team Writers");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                ls = p.getListField("Targeted countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                if (!task.getStringField("Work Package ID").isEmpty()) {
                                                    valueToWrite.add(task.getStringField("Work Package ID").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 2").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Manager 3").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                                    valueToWrite.add(task.getStringField("workPackagePriority").replace("\n", ",").replace(";", ","));
                                                    ls = task.getListField("workPackageUserWriter");
                                                    text = "";
                                                    for (String l : ls) {
                                                        if (!text.equals("")) {
                                                            text += ",";
                                                        }
                                                        text += l;
                                                    }
                                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                } else {
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                    valueToWrite.add("");
                                                }

                                                ls = task.getListField("TO Year 5 Countries");
                                                text = "";
                                                for (String l : ls) {
                                                    if (!text.equals("")) {
                                                        text += ",";
                                                    }
                                                    text += l;
                                                }
                                                valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 k€")).replace(".", ","));
                                                valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 Market value")).replace(".", ","));
                                                valueToWrite.add(task.getStringField("TO Year 5 Zones").replace("\n", ",").replace(";", ","));

                                                CSVUtils.writeLine(writer, valueToWrite);
                                                valueToWrite = new ArrayList<>();
                                                cDebut.add(Calendar.DATE, 7);
                                                cDebutFin.add(Calendar.DATE, 7);
                                            }
                                        }
                                    }
                                }
                                if (task.getBooleanField("milestones")) {
                                    //**************************************************************************************\\
                                    Logger.info("  Milestones '" + task.getStringField("name") + "':");
                                    //**************************************************************************************\\
                                    cDebut.setTime(task.getDateField("Start"));
                                    cDebutFin.setTime(task.getDateField("Start"));
                                    cDebutFin.add(Calendar.DATE, 6);
                                    cFin.setTime(task.getDateField("Finish"));
                                    valueToWrite.add(sdf.format(cDebut.getTime()));
                                    valueToWrite.add(sdf.format(cDebutFin.getTime()));
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.YEAR)));
                                    valueToWrite.add((cDebut.get(Calendar.MONTH) >= Calendar.JANUARY && cDebut.get(Calendar.MONTH) <= Calendar.MARCH) ? "Q1" : (cDebut.get(Calendar.MONTH) >= Calendar.APRIL && cDebut.get(Calendar.MONTH) <= Calendar.JUNE) ? "Q2" : (cDebut.get(Calendar.MONTH) >= Calendar.JULY && cDebut.get(Calendar.MONTH) <= Calendar.SEPTEMBER) ? "Q3" : "Q4");
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.MONTH) + 1));
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.WEEK_OF_YEAR) == 53 ? 1 : cDebut.get(Calendar.WEEK_OF_YEAR) + 1));
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");//valueToWrite.add(transaction.getStringField("Organization"));
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add(String.valueOf(task.getIntField("#")));

                                    List<String> ls = task.getListField("Tag Zone");
                                    String text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = task.getListField("Tag Project Scope");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(task.getStringField("Tag PMO Scope").replace("\n", ",").replace(";", ",")); //Tag Portfolio Scope?

                                    ls = task.getListField("Tag Department Scope");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = task.getListField("Tag Countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(task.getBooleanField("Tag Major Country in Zone") ? "true" : "false");
                                    valueToWrite.add(task.getStringField("Tag Chapter").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("PHARMA - Study Code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("ID").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getBooleanField("milestones") ? "Milestones" : "Task");
                                    valueToWrite.add(task.getStringField("Hierarchy N-1").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-2").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-3").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-4").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-5").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-6").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-7").replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("ID").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Name").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("priority").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Strategic Target Group");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Product code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project Code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project Type").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Type").replace("\n", ",").replace(";", ",")); //Project Type / Category identique ?
                                    valueToWrite.add(p.getStringField("Franchise").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Target species");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Therapeutic axis");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Intended indication(s)").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Innovation level").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Targeted zones");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getBooleanField("Strategic Product") ? "true" : "false");

                                    ls = p.getListField("API (active ingredient)");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Dosage form").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Portfolio Folder").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project owner").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(sdf.format(p.getDateField("Modified Date")));
                                    valueToWrite.add(sdf.format(p.getDateField("Published Date")));
                                    valueToWrite.add(p.getStringField("Coordinator").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Workflow State").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Frise Phase project").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project phase").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Offer Conception manager").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Patent team member").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("POC process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Galenic process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Industrialization process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Safety process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Dose-Indication process pilot").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Reg Affairs team leaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Marketing team member").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Supply Chain team member").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Performing Organizations");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Owning Organization").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("usersReaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("usersWriters");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("organizationReaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("organizationWriters");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Core Team Readers");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Core Team Writers");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Targeted countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    if (!task.getStringField("Work Package ID").isEmpty()) {
                                        valueToWrite.add(task.getStringField("Work Package ID").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 2").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 3").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("workPackagePriority").replace("\n", ",").replace(";", ","));
                                        ls = task.getListField("workPackageUserWriter");
                                        text = "";
                                        for (String l : ls) {
                                            if (!text.equals("")) {
                                                text += ",";
                                            }
                                            text += l;
                                        }
                                        valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    } else {
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                    }

                                    ls = task.getListField("TO Year 5 Countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 k€")).replace(".", ","));
                                    valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 Market value")).replace(".", ","));
                                    valueToWrite.add(task.getStringField("TO Year 5 Zones").replace("\n", ",").replace(";", ","));

                                    CSVUtils.writeLine(writer, valueToWrite);
                                    valueToWrite = new ArrayList<>();
                                }
                                if (task.getBooleanField("isWorkPackage")) {
                                    //**************************************************************************************\\
                                    Logger.info("  Work Package '" + task.getStringField("name") + "':");
                                    //**************************************************************************************\\
                                    cDebut.setTime(task.getDateField("Start"));
                                    cDebutFin.setTime(task.getDateField("Start"));
                                    cDebutFin.add(Calendar.DATE, 6);
                                    cFin.setTime(task.getDateField("Finish"));
                                    valueToWrite.add(sdf.format(cDebut.getTime()));
                                    valueToWrite.add(sdf.format(cDebutFin.getTime()));
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.YEAR)));
                                    valueToWrite.add((cDebut.get(Calendar.MONTH) >= Calendar.JANUARY && cDebut.get(Calendar.MONTH) <= Calendar.MARCH) ? "Q1" : (cDebut.get(Calendar.MONTH) >= Calendar.APRIL && cDebut.get(Calendar.MONTH) <= Calendar.JUNE) ? "Q2" : (cDebut.get(Calendar.MONTH) >= Calendar.JULY && cDebut.get(Calendar.MONTH) <= Calendar.SEPTEMBER) ? "Q3" : "Q4");
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.MONTH) + 1));
                                    valueToWrite.add(String.valueOf(cDebut.get(Calendar.WEEK_OF_YEAR) == 53 ? 1 : cDebut.get(Calendar.WEEK_OF_YEAR) + 1));
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");//valueToWrite.add(transaction.getStringField("Organization"));
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add("");
                                    valueToWrite.add(String.valueOf(task.getIntField("#")));

                                    List<String> ls = task.getListField("Tag Zone");
                                    String text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = task.getListField("Tag Project Scope");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(task.getStringField("Tag PMO Scope").replace("\n", ",").replace(";", ",")); //Tag Portfolio Scope?

                                    ls = task.getListField("Tag Department Scope");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = task.getListField("Tag Countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(task.getBooleanField("Tag Major Country in Zone") ? "true" : "false");
                                    valueToWrite.add(task.getStringField("Tag Chapter").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("PHARMA - Study Code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("ID").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getBooleanField("milestones") ? "Milestones" : "Task");
                                    valueToWrite.add(task.getStringField("Hierarchy N-1").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-2").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-3").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-4").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-5").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-6").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(task.getStringField("Hierarchy N-7").replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("ID").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Name").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("priority").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Strategic Target Group");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Product code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project Code").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project Type").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Type").replace("\n", ",").replace(";", ",")); //Project Type / Category identique ?
                                    valueToWrite.add(p.getStringField("Franchise").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Target species");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Therapeutic axis");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Intended indication(s)").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Innovation level").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Targeted zones");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getBooleanField("Strategic Product") ? "true" : "false");

                                    ls = p.getListField("API (active ingredient)");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Dosage form").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Portfolio Folder").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project owner").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(sdf.format(p.getDateField("Modified Date")));
                                    valueToWrite.add(sdf.format(p.getDateField("Published Date")));
                                    valueToWrite.add(p.getStringField("Coordinator").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Workflow State").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Frise Phase project").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Project phase").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Offer Conception manager").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Patent team member").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("POC process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Galenic process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Industrialization process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Safety process pilot").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Dose-Indication process pilot").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Reg Affairs team leaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));

                                    valueToWrite.add(p.getStringField("Marketing team member").replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Supply Chain team member").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("Performing Organizations");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(p.getStringField("Owning Organization").replace("\n", ",").replace(";", ","));

                                    ls = p.getListField("usersReaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("usersWriters");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("organizationReaders");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("organizationWriters");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Core Team Readers");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Core Team Writers");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    ls = p.getListField("Targeted countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    if (!task.getStringField("Work Package ID").isEmpty()) {
                                        valueToWrite.add(task.getStringField("Work Package ID").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 1").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 2").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Manager 3").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("Name").replace("\n", ",").replace(";", ","));
                                        valueToWrite.add(task.getStringField("workPackagePriority").replace("\n", ",").replace(";", ","));
                                        ls = task.getListField("workPackageUserWriter");
                                        text = "";
                                        for (String l : ls) {
                                            if (!text.equals("")) {
                                                text += ",";
                                            }
                                            text += l;
                                        }
                                        valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    } else {
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                        valueToWrite.add("");
                                    }

                                    ls = task.getListField("TO Year 5 Countries");
                                    text = "";
                                    for (String l : ls) {
                                        if (!text.equals("")) {
                                            text += ",";
                                        }
                                        text += l;
                                    }
                                    valueToWrite.add(text.replace("\n", ",").replace(";", ","));
                                    valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 k€")).replace(".", ","));
                                    valueToWrite.add(nf.format(task.getDoubleField("TO Year 5 Market value")).replace(".", ","));
                                    valueToWrite.add(task.getStringField("TO Year 5 Zones").replace("\n", ",").replace(";", ","));

                                    CSVUtils.writeLine(writer, valueToWrite);
                                    valueToWrite = new ArrayList<>();
                                }
                            }
                        }
                    }//
                } catch (PSException ex) {
                    Logger.error(ex);
                } finally {
                    p.close();
                }

            }
            writer.flush();
            writer.close();
        } catch (PSException ex) {
            Logger.error(ex);
        } catch (Exception ex) {
            Logger.error(ex);
            System.exit(1);
        }
    }
}
