package ext.mpxj;  
  
import com.jacob.activeX.ActiveXComponent;  
import com.jacob.com.Dispatch;  
import com.jacob.com.Variant;  
  
import ext.mpxj.dao.Assignment;  
import ext.mpxj.dao.Resource;  
import ext.mpxj.dao.Task;  
  
public class MPPFileOperation {  
      
    private ActiveXComponent msProjApp = null;  
    private boolean isSave = true;  
    private Dispatch activeProject = null;  
    public ActiveXComponent getMsProjApp() {  
        return msProjApp;  
    }  
  
    public void setMsProjApp(ActiveXComponent msProjApp) {  
        this.msProjApp = msProjApp;  
    }  
  
    public boolean isSave() {  
        return isSave;  
    }  
  
    public void setSave(boolean isSave) {  
        this.isSave = isSave;  
    }  
  
    public Dispatch getActiveProject() {  
        return activeProject;  
    }  
  
    public void setActiveProject(Dispatch activeProject) {  
        this.activeProject = activeProject;  
    }  
  
    public Dispatch getTasks() {  
        return tasks;  
    }  
  
    public void setTasks(Dispatch tasks) {  
        this.tasks = tasks;  
    }  
  
    public Dispatch getResources() {  
        return resources;  
    }  
  
    public void setResources(Dispatch resources) {  
        this.resources = resources;  
    }  
  
  
    private Dispatch tasks = null;  
    private Dispatch resources = null;  
      
      
      
      
    /** 
     * Create MPP File, by exising one or new one 
     * @param path 
     * @param isSave 
     */  
    public MPPFileOperation(String path, boolean isSave) {  
        msProjApp = new ActiveXComponent("MSPROJECT.Application");  
        this.isSave = isSave;  
        if (isSave) {  
            Dispatch.call(msProjApp, "FileNew", Dispatch.Method);  
            Dispatch.call(msProjApp, "SaveAs", new Variant(path));  
        } else {  
            Dispatch.invoke(msProjApp, "FileOpen", Dispatch.Method, new Object[]{path}, new int[]{1});  
        }  
        // Select current Active Project  
        activeProject = msProjApp.getProperty("ActiveProject").toDispatch();  
        tasks = Dispatch.get(activeProject, "Tasks").toDispatch();  
        resources = Dispatch.get(activeProject, "Resources").toDispatch();  
        isSave = false;  
    }  
      
    /** 
     * Do mpp file save 
     */  
    public void save() {  
        Dispatch.call(msProjApp, "FileSave");  
    }  
      
    /** 
     * close mpp, if exist one, update, otherwise save a new one 
     */  
    public void close() {  
        if (isSave) {  
            save();  
        } else {  
            save();  
        }  
        Dispatch.call(msProjApp, "Quit");  
    }  
      
    /** 
     * Create task according to Name 
     * @param taskName 
     * @return 
     */  
    public Task createTask(String taskName) {  
        Task task = null;  
        if (tasks != null) {  
            task = new Task(Dispatch.invoke(tasks, "Add", Dispatch.Method, new Object[]{taskName}, new int[]{1}).toDispatch(), msProjApp);  
        }  
        return task;  
    }  
      
    /** 
     * Create Summary Task over subtask 
     * @param taskName 
     * @param childTasks 
     * @return 
     * @throws Exception  
     */  
    public Task createSummaryTask(String taskName, String nextTaskName, String... childTasks) throws Exception {  
        Task sumTask = null;  
        if (nextTaskName == null || "".equals(nextTaskName)) {  
            throw new Exception("Next Taskname can't be null or empty string!");  
        }  
        if (tasks != null) {  
            Dispatch sumTaskDispatch = Dispatch.invoke(tasks, "Add", Dispatch.Method, new Object[]{taskName}, new int[]{1}).toDispatch();  
            for (int i = 0; i < childTasks.length;i++) {  
                Dispatch childTaskDispatch = Dispatch.invoke(tasks, "Add", Dispatch.Method, new Object[]{childTasks[i]}, new int[]{1}).toDispatch();  
                if (i == 0) {  
                    Dispatch.invoke(childTaskDispatch, "OutlineIndent", Dispatch.Method, new Object[]{}, new int[]{1});  
                }  
            }  
              
            sumTask = new Task(sumTaskDispatch, msProjApp);  
            Dispatch newTaskDispatch = Dispatch.invoke(tasks, "Add", Dispatch.Method, new Object[]{nextTaskName}, new int[]{1}).toDispatch();  
            Dispatch.invoke(newTaskDispatch, "OutlineOutdent", Dispatch.Method, new Object[]{}, new int[]{1});  
        }  
        return sumTask;  
    }  
      
    /** 
     * Iterate all task and find the task equal to given name 
     * @param taskName 
     * @return 
     */  
    public Task findTaskByName(String taskName) {  
        Task task = null;  
        int count = Dispatch.get(tasks, "Count").getInt();  
        for (int i = 0; i < count; i++) {  
            Dispatch curTask = Dispatch.invoke(tasks, "Item", Dispatch.Get, new Object[]{i+1}, new int[]{1}).toDispatch();  
            String curTaskName = Dispatch.get(curTask, "Name").getString();  
            if (taskName.equals(curTaskName)) {  
                task = new Task(curTask, msProjApp);;  
                break;  
            }  
        }  
        return task;  
    }  
      
    /** 
     * Create resource by given name 
     * @param resourceName 
     * @return 
     */  
    public Resource createResource(String resourceName) {  
        Resource resource = null;  
        Dispatch sumTaskDispatch = Dispatch.invoke(resources, "Add", Dispatch.Method, new Object[]{resourceName}, new int[]{1}).toDispatch();  
        resource = new Resource(sumTaskDispatch, msProjApp);  
        return resource;  
    }  
      
    /** 
     * Iterate all resource and find the resource matches 
     * @param resourceName 
     * @return 
     */  
    public Resource findResourceByName(String resourceName) {  
        Resource resource = null;  
        int count = Dispatch.get(resources, "Count").getInt();  
        for (int i = 0; i < count; i++) {  
            Dispatch curResource = Dispatch.invoke(resources, "Item", Dispatch.Get, new Object[]{i+1}, new int[]{1}).toDispatch();  
            String curResourceName = Dispatch.get(curResource, "Name").getString();  
            if (resourceName.equals(curResourceName)) {  
                resource = new Resource(curResource, msProjApp);  
                break;  
            }  
        }  
        return resource;  
    }  
      
    /** 
     * Create assignments combines task and resource 
     * @param task 
     * @param resourceName 
     * @return 
     * @throws Exception  
     */  
    public Assignment createAssignment(Task task, Resource resc) throws Exception {  
        return task.addAssignment(resc);  
    }  
      
    /** 
     * find assignment by resourceName and given task 
     * @param task 
     * @param resourceName 
     * @return 
     */  
    public Assignment findAssignmentByName(Task task, String resourceName) {  
        Assignment assignment = null;  
        Dispatch assignmentsDispatch = Dispatch.get(task.getTask(), "Assignments").toDispatch();  
        int count = Dispatch.get(assignmentsDispatch, "Count").getInt();  
        for (int i = 0;i < count;i++) {  
            Dispatch curAssignment = Dispatch.invoke(assignmentsDispatch, "Item", Dispatch.Method, new Object[]{i+1}, new int[]{1}).toDispatch();  
            String curAssignmentName = Dispatch.get(curAssignment, "Name").getString();  
            if (resourceName.equals(curAssignmentName)) {  
                assignment = new Assignment(task, msProjApp, curAssignment);  
                break;  
            }  
        }  
        return assignment;  
    }  
      
      
    public void promoteToSummaryTask(Task task, String... subTasks) {  
        boolean isSummary = Dispatch.get(task.getTask(), "Summary").getBoolean();  
        if (!isSummary) {  
            for (int i=subTasks.length;i>0;i--) {  
                String subTask = subTasks[i-1];  
                Dispatch subTaskDispatch = Dispatch.invoke(tasks, "Add", Dispatch.Method, new Object[]{subTask, task.getTaskID() + 1}, new int[]{1}).getDispatch();  
                if (i == subTasks.length) {  
                    Dispatch.invoke(subTaskDispatch, "OutlineIndent", Dispatch.Method, new Object[]{}, new int[]{1});  
                }  
            }  
        } else {  
              
        }  
    }  
      
      
    public void addCustomField(int tableIdx, String fieldName, int pjField) {  
        Dispatch tablesDispatch = Dispatch.get(activeProject, "TaskTables").getDispatch();  
        Dispatch tableDispatch = Dispatch.invoke(tablesDispatch, "Item", Dispatch.Get, new Object[]{tableIdx}, new int[]{1}).toDispatch();  
        System.out.println(Dispatch.get(tableDispatch, "Name").getString());  
          
        Dispatch tableFieldsDispatch = Dispatch.invoke(tableDispatch, "TableFields", Dispatch.Get, new Object[]{}, new int[]{1}).toDispatch();  
        Variant specifiedVar = new Variant();  
        specifiedVar.putString(fieldName);  
        specifiedVar.putNoParam();  
        Dispatch.invoke(tableFieldsDispatch, "Add", Dispatch.Method, new Object[]{pjField,specifiedVar,specifiedVar,fieldName}, new int[]{1});  
    }  
  
}  
