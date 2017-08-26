package ext.project.ms.msproject;
import com.jacob.activeX.ActiveXComponent;  
import com.jacob.com.Dispatch;  
  
public class Assignment implements IProjectData {  
    private static final long serialVersionUID = 8518032497041538417L;  
    private Task task;  
    private ActiveXComponent msProjApp;  
    public Task getTask() {  
        return task;  
    }  
  
    public void setTask(Task task) {  
        this.task = task;  
    }  
  
    public ActiveXComponent getMsProjApp() {  
        return msProjApp;  
    }  
  
    public void setMsProjApp(ActiveXComponent msProjApp) {  
        this.msProjApp = msProjApp;  
    }  
  
    public Dispatch getAssignment() {  
        return assignment;  
    }  
  
    public void setAssignment(Dispatch assignment) {  
        this.assignment = assignment;  
    }  
  
    private Dispatch assignment;  
    public Assignment(Task task, ActiveXComponent msProjApp, Dispatch assignment) {  
        this.task = task;  
        this.msProjApp = msProjApp;  
        this.assignment = assignment;  
    }  
      
    public void update(String key, Object value) {  
        Dispatch.put(assignment, key, value);  
    }  
  
    public void delete() {  
        Dispatch.invoke(assignment, "Delete", Dispatch.Method, new Object[]{}, new int[]{1}).toDispatch();  
    }  
}  