package thoan_excel;

public class FacebookUser {
	String id;
	String link;
	String name;
	
	
	
	public FacebookUser(String name, String id, String link) {
		super();
		this.id = id;
		this.link = link;
		this.name = name;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getLink() {
		return link;
	}
	public void setLink(String link) {
		this.link = link;
	}
	
}
