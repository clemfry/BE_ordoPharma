import jxl.Workbook;
import jxl.format.ScriptStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import org.chocosolver.solver.Model;
import org.chocosolver.solver.Solution;
import org.chocosolver.solver.search.strategy.Search;
import org.chocosolver.solver.variables.BoolVar;
import org.chocosolver.solver.variables.IntVar;
import org.chocosolver.solver.variables.Task;
import org.chocosolver.util.tools.ArrayUtils;

import java.io.File;
import java.io.IOException;


public class OrdoPB {

    public static void main(String[] args) {
        WritableWorkbook workbook = null;
        boolean q4 = true; // For the question #4, engines #1 and #2 can't work simultaneously
        boolean q6 = false; // For the question #6, there are 2 engines for process 1

        try {
            Model model = new Model("Ordo Pharma");
            long seed = 18081981;


            // Definition of the problem parameters
            int[][] T = new int[][]{ // duration of the process of the 2 steps of each of the 3 products [duration of step 1, duration of step 2]
                    {10,30},
                    {30,70},
                    {20,30}
            };

            int[][] orderList = new int[][]{ // orders [productType, qty, deadline]
                    {0, 5, 700},
                    {1, 2, 2500},
                    {0, 4, 2500},
                    {2, 7, 2500},
                    {2, 3, 2500},
                    {1, 1, 2500},
                    {1, 6, 2500},
                    {0, 9, 2500},
            };

            // here we consider 4 fabrication orders (FO), each composed by 2 tasks and 4 specific engines, of unitary capacity,
            // ensuring the production
            int maxTime = 2500;
            int nbO = orderList.length; // number of FO
            int nbTO = 2; // number of tasks for each FO
            int nbT = nbO * nbTO; // total number of tasks
            int nbR; // number of engines/resources of unitary capacity
            if (q6 == false){nbR = 4;} else { nbR = 5;}


            // engines attributes declaration
            IntVar cr = model.intVar("capacity", nbR, nbR); // total capacity of all the engines/resources
            IntVar[] crr = model.intVarArray("unitaryCapa", nbR, 1, 1); // unitary capacity of each engine/resource

            // tasks attributes declaration (there are nbT tasks in total)
            IntVar[] s = model.intVarArray("s", nbT, 0, maxTime); // starts
            IntVar[] e = model.intVarArray("e", nbT, 0, maxTime); // ends
            IntVar[] d = new IntVar[nbT];
            for (int i = 0; i < nbO; i++) {
                int productType = orderList[i][0]; // type of product for order i
                int qty = orderList[i][1]; // quantity of product for order i
                int dPhase1 = T[productType][0]*qty; // duration phase 1
                int dPhase2 = T[productType][1]*qty; // duration phase 2
                d[2*i] = model.intVar("d", dPhase1, dPhase1); // durations
                d[2*i +1] = model.intVar("d", dPhase2, dPhase2); // durations
            }
            IntVar[] c = model.intVarArray("c", nbT, 1, 1); // unitary resources consumption of each tasks

            // tasks creation
            Task[] tasks = new Task[nbT];
            System.out.println("*************************");
            System.out.println("Considering the following set of tasks:");
            for (int i = 0; i < nbT; i++) {
                tasks[i] = new Task(s[i], d[i], e[i]); // container modeling a task, ensure : s[i] + d[i] = e[i]
                System.out.println(tasks[i]);
            }
            System.out.println("subjected to s[i] + d[i] = e[i]");
            System.out.println("*************************");

            // engine execution / for each task, true iff the engine i executes task j
            BoolVar[][] ex = new BoolVar[nbR][nbT];
            for (int i = 0; i < nbR; i++) {
                for (int j = 0; j < nbT; j++) {
                    ex[i][j] = model.boolVar("ex" + i + "," + j);

                    // Some engines can only execute specific tasks
                    int orderIndex = j/2;
                    if (2*orderIndex == j ){  // if the task is phase one (index is even)
                        if ( (i!=0) && (i !=4) ) {
                            model.arithm(ex[i][j], "=", 0).post();// only the engine #0 can execute this task (or engine #4 if q6 = true)
                        }
                    } else {    // if the task is phase two (index is odd)
                        int productType = orderList[orderIndex][0]; // type of product for order orderIndex
                        if (i == productType + 1){
                            model.arithm(ex[i][j], "=", 1).post(); // only the engine #productType+1 can execute this task
                        } else {
                            model.arithm(ex[i][j], "=", 0).post();
                        }
                    }
                }
            }

            // grouping tasks into the FOs
            Task[][] orders = new Task[nbO][nbTO];
            System.out.println("*************************");
            System.out.println("Considering the following set of fabrication orders:");
            for (int i = 0; i < nbT; ) {
                int o = (i / nbTO);
                System.out.print("Fabrication Order " + o + " :: ");
                for (int j = 0; j < nbTO; j++) {
                    orders[o][j] = tasks[i++];
                    if (j >= 1) {
                        System.out.print(", Task[" + orders[o][j].getStart().getName() + ";" + orders[o][j].getEnd().getName() + "]");
                    } else {
                        System.out.print("Task[" + orders[o][0].getStart().getName() + ";" + orders[o][0].getEnd().getName() + "]");
                    }
                }
                System.out.println();
            }
            System.out.println("*************************");

            // cumulative constraint ensures that, at each point of time, the total resources consumption
            // of tasks planed does not exceed the total capacity available
            model.cumulative(tasks, c, cr).post();

            // for each order, task i have to finish before task i+1 starts
            System.out.println("*************************");
            System.out.println("Tasks inside each FO are ordered");
            for (int k = 0; k < nbO; k++) {
                System.out.print("FO " + k + " :: " + orders[k][0].getStart().getName() + " - " + orders[k][0].getEnd().getName());
                for (int i = 0; i < nbTO - 1; i++) {
                    model.arithm(orders[k][i].getEnd(), "<=", orders[k][i + 1].getStart()).post();
                    System.out.print(" <= " + orders[k][i + 1].getStart().getName() + " - " + orders[k][i + 1].getEnd().getName());
                }
                System.out.println();
            }
            System.out.println("*************************");

            // for each order, task must finish before the given deadline
            System.out.println("*************************");
            System.out.println("FOs must end before deadline");
            for (int k = 0; k < nbO; k++) {
                System.out.print("FO " + k + " :: " + orders[k][1].getStart().getName() + " - " + orders[k][1].getEnd().getName() + " <= " + orderList[k][2]);
                model.arithm(orders[k][1].getEnd(), "<=", orderList[k][2]).post();
                System.out.println();
            }
            System.out.println("*************************");

            // tasks attributes declaration per engine
            Task[][] taskR = new Task[nbR][nbT];
            IntVar[][] se = new IntVar[nbR][nbT];
            IntVar[][] ee = new IntVar[nbR][nbT];
            IntVar[][] de = new IntVar[nbR][nbT];
            IntVar[][] ce = new IntVar[nbR][nbT];
            for (int i = 0; i < nbR; i++) {
                for (int j = 0; j < nbT; j++) {
                    se[i][j] = model.intVar("se" + i + "," + j, 0, maxTime); // start date of engine i for task j
                    ee[i][j] = model.intVar("ee" + i + "," + j, 0, maxTime); // end date of engine i for task j
                    de[i][j] = model.intVar("de" + i + "," + j, 0, maxTime); // duration of work of engine i for task j
                    ce[i][j] = model.intVar("ce" + i + "," + j, 0, 1); // nb fo resources required by this task (=1 or 0 if engine i makes task j)
                    model.arithm(se[i][j], "=", ex[i][j], "*", s[j]).post(); // se[i][j] = ex[i][j] * s[j]
                    model.arithm(ee[i][j], "=", ex[i][j], "*", e[j]).post(); // ee[i][j] = ex[i][j] * e[j]
                    model.arithm(de[i][j], "=", ex[i][j], "*", d[j]).post(); // de[i][j] = ex[i][j] * d[j]
                    model.arithm(ce[i][j], "=", ex[i][j], "*", c[j]).post(); // ce[i][j] = ex[i][j] * c[j]
                    taskR[i][j] = new Task(se[i][j], de[i][j], ee[i][j]);
                }
            }

            System.out.println("*************************");
            // one cumulative for each engine ensuring engine capacity
            System.out.println("Each engine/resource cannot execute more than one task at a time");
            for (int i = 0; i < nbR; i++) {
                model.cumulative(taskR[i], ce[i], crr[i]).post();
                System.out.println("Engine-" + i + " :: cumulative(all possible tasks executed on this engine)");
            }

            if (q4== true){
                System.out.println("*************************");
                // two of the engines cannot work simultaneously (#1 and #2 here)
                int e1 = 1; int e2 = 2; // engines that can't work simultaneously
                System.out.println("engine-" + e1 + " and engine-" + e2 + " can't work simultaneaously");

                int concatLength = taskR[e1].length + taskR[e2].length;
                Task[] taskRConcat = new Task[concatLength];
                IntVar[] ceConcat = new IntVar[concatLength];
                int pos = 0;
                for (Task t : taskR[e1]){
                    taskRConcat[pos] = t;
                    ceConcat[pos] = ce[e1][pos];
                    pos ++;
                }
                for (Task t : taskR[e2]){
                    taskRConcat[pos] = t;
                    ceConcat[pos] = ce[e2][pos - taskR[e1].length];
                    pos ++;
                }
                model.cumulative(taskRConcat, ceConcat, crr[e1]).post();
            }

            // a task is executed by exactly one engine
            BoolVar[][] revTask = ArrayUtils.transpose(ex);
            for (int j = 0; j < nbT; j++) {
                model.sum(revTask[j], "=", 1).post();
            }
            System.out.println("*************************");

            // configuration of the search strategy
            int search = 3;
            workbook = null;
            switch(search){
                case 1:
                    /* creating new workbook and writting */
                    workbook = Workbook.createWorkbook(new File("ordoResultLB.xls"));
                    // choose the tasks in the input order and instantiate their starting date as soon as possible
                    // and next, instantiate each task to the first engine/resource available
                    model.getSolver().setSearch(Search.inputOrderLBSearch(s),Search.inputOrderLBSearch(ArrayUtils.flatten(ex)));
                    break;
                case 2:
                    /* creating new workbook and writting */
                    workbook = Workbook.createWorkbook(new File("ordoResultMinDom.xls"));
                    // choose first the task with the minimum domain size and instantiate it to the starting date as soon as possible
                    // and next, instantiate each task to the first engine/resource available
                    model.getSolver().setSearch(Search.minDomLBSearch(s),Search.inputOrderLBSearch(ArrayUtils.flatten(ex)));
                    break;
                default:
                    /* creating new workbook and writting */
                    workbook = Workbook.createWorkbook(new File("ordoResultDomWdeg.xls"));
                    // Finding the most promising task (a bit copmplex - no details)
                    // and next, random choice to instantiate the tasks to the engines/resources
                    model.getSolver().setSearch(Search.domOverWDegSearch(s),Search.randomSearch(ArrayUtils.flatten(ex),seed));
            }


            // starting resolution to find one solution if it exists
            Solution solution = model.getSolver().findSolution();
            if (solution != null) {
                System.out.println("\n\n");
                System.out.println("*************************");
                System.out.println("solution found: ");
                for (int i = 0; i < nbR; i++) {
                    System.out.print("Engine-" + i + " execution schedule :: ");

                    for (int j = 0; j < nbT; j++) {
                        if (ex[i][j].isInstantiatedTo(1)) {
                            System.out.print("T" + j + "-[" + s[j].getValue() + "," + e[j].getValue() + "] ");
                        }
                    }
                    System.out.println();
                }
                System.out.println("*************************");

                System.out.println("*************************");
                System.out.println("Pretty print results");
                /* creating new sheet and writting */
                WritableSheet sheet = workbook.createSheet("OrdoResultat", 0);
                /* creating format */
                //Crée le format d’une cellule
                WritableFont font = new WritableFont(WritableFont.ARIAL, 11, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLACK, ScriptStyle.NORMAL_SCRIPT);
                WritableCellFormat format = new WritableCellFormat(font);
                /* Creating text data */
                Label time = new Label(0, 0, "Time slots", format);
                sheet.addCell(time);
                for (int i = 0; i < (maxTime/10); i++) {
                    Label label = new Label(i + 1, 0, "[" + i * 10 + "," + (i + 1) * 10 + "]", format);
                    sheet.addCell(label);
                }
                for (int i = 0; i < nbR; i++) {
                    Label label = new Label(0, i + 1, "Engine-" + i, format);
                    sheet.addCell(label);
                }
                /* field creation by engines per tasks */
                WritableFont font2 = new WritableFont(WritableFont.ARIAL, 11, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.WHITE, ScriptStyle.NORMAL_SCRIPT);
                int colour = 1;
                for (int i = 0; i < nbR; i++) {
                    for (int j = 0; j < nbT; j++) {
                        if (ex[i][j].isInstantiatedTo(1)) {
                            System.out.println("task" + j + " : [" + s[j].getValue() + "," + e[j].getValue() + "]");
                            int colStart = (s[j].getValue() / 10) + 1;
                            int colEnd = (e[j].getValue() / 10) + 1;
                            for (int r = colStart; r < colEnd; r++) {
                                switch (colour) {
                                    case 1:
                                        WritableCellFormat f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.BLUE);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 2:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.OLIVE_GREEN);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 3:
                                        f = new WritableCellFormat(font);
                                        f.setBackground(Colour.PALE_BLUE);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 4:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.BLUE_GREY);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 5:
                                        f = new WritableCellFormat(font);
                                        f.setBackground(Colour.AQUA);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 6:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.BROWN);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 7:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.CORAL);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 8:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.DARK_BLUE);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 9:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.DARK_GREEN);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 10:
                                        f = new WritableCellFormat(font);
                                        f.setBackground(Colour.LIGHT_BLUE);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 11:
                                        f = new WritableCellFormat(font);
                                        f.setBackground(Colour.LIGHT_GREEN);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    case 12:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.OCEAN_BLUE);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                        break;
                                    default:
                                        f = new WritableCellFormat(font2);
                                        f.setBackground(Colour.YELLOW);
                                        sheet.addCell(new Label(r, i + 1, "Task" + j, f));
                                }
                            }
                            colour++;
                        }
                    }
                }
                /* writting the workbook */
                workbook.write();
                System.out.println("*************************");
            } else {
                System.out.println("no solution");
            }
        } catch (IOException | RowsExceededException e0) {
            e0.printStackTrace();
        } catch (WriteException e1) {
            e1.printStackTrace();
        } finally {
            if (workbook != null) {
                /* closing workbook */
                try {
                    workbook.close();
                } catch (WriteException | IOException e2) {
                    e2.printStackTrace();
                }
            }
        }
    }

}
