package ext.project.ms.msproject;
import com.jacob.activeX.ActiveXComponent;  
import com.jacob.com.Dispatch;  
  
public class Resource implements IProjectData {  
    private static final long serialVersionUID = 295721362536771540L;  
    private Dispatch resource;  
    private ActiveXComponent msProjApp;  
  
    public Dispatch getResource() {  
        return resource;  
    }  
  
    public void setResource(Dispatch resource) {  
        this.resource = resource;  
    }  
  
    public ActiveXComponent getMsProjApp() {  
        return msProjApp;  
    }  
  
    public void setMsProjApp(ActiveXComponent msProjApp) {  
        this.msProjApp = msProjApp;  
    }  
  
    public Resource(Dispatch resource, ActiveXComponent msProjApp) {  
        this.resource = resource;  
        this.msProjApp = msProjApp;  
    }  
      
    public void update(String colName, Object colVal) {  
        int fieldID = Dispatch.invoke(msProjApp, "FieldNameToFieldConstant", Dispatch.Method, new Object[]{colName, COMObjectTypeConstant.pjResource}, new int[]{1}).getInt();  
        Dispatch.invoke(resource, "SetField", Dispatch.Method, new Object[]{fieldID, colVal}, new int[]{1});  
    }  
      
    public void delete() {  
        Dispatch.invoke(resource, "Delete", Dispatch.Method, new Object[]{}, new int[]{1});  
    }  
      
}  