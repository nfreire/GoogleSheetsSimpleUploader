package apiclient.google.sheets;

public class ImageValue {
	String imageUrl;

	public ImageValue(String imageUrl) {
		this.imageUrl = imageUrl;
	}

	@Override
	public String toString() {
		return imageUrl;
	}
	
}
