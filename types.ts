export interface ProcessResponse {
  images: GeneratedImage[];
  zip_url: string;
  message: string;
}

export interface GeneratedImage {
  project_name: string;
  image_url: string;
  filename: string;
}

export interface ApiError {
  detail: string;
}
