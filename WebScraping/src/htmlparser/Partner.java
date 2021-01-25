package htmlparser;

public class Partner {
	String partnerAccountId;
	String caAccountId;
	String id;
	String website;
	String email;
	String temp;
	
	public Partner(String partnerAccountId, String caAccountId, String id) {
		super();
		this.partnerAccountId = partnerAccountId;
		this.caAccountId = caAccountId;
		this.id = id;
	}
	
	public String getTemp() {
		return temp;
	}

	public void setTemp(String temp) {
		this.temp = temp;
	}

	public String getWebsite() {
		return website;
	}

	public void setWebsite(String website) {
		this.website = website;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getPartnerAccountId() {
		return partnerAccountId;
	}
	public void setPartnerAccountId(String partnerAccountId) {
		this.partnerAccountId = partnerAccountId;
	}
	public String getCaAccountId() {
		return caAccountId;
	}
	public void setCaAccountId(String caAccountId) {
		this.caAccountId = caAccountId;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	
	
}
