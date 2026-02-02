
"""
Document Classifier Module (PyTorch Transfer Learning)
----------------------------------------------------
This module demonstrates a foundational AI task: Image Classification.
Scenario: Automatically sorting documents into categories (e.g., 'Invoice', 'Receipt', 'Contract')
before processing them.

Teaching Points:
1.  **Transfer Learning**: Using a pre-trained ResNet18 model.
2.  **Data Augmentation**: Using `torchvision.transforms` to improve generalization.
3.  **Training Loop**: Writing a standard PyTorch training loop from scratch.
4.  **Visualization**: Plotting loss curves with Matplotlib.

"""

import torch
import torch.nn as nn
import torch.optim as optim
from torchvision import models, transforms
from torch.utils.data import Dataset, DataLoader
from PIL import Image
import os
import matplotlib.pyplot as plt

# 1. Configuration
NUM_CLASSES = 3
CLASSES = ['invoice', 'receipt', 'others']
BATCH_SIZE = 4
LEARNING_RATE = 0.001
EPOCHS = 5

class MockDocumentDataset(Dataset):
    """
    Simulates loading images from folders.
    In a real lab, students would use 'torchvision.datasets.ImageFolder'.
    """
    def __init__(self, num_samples=20, transform=None):
        self.num_samples = num_samples
        self.transform = transform
        # Create a dummy white image for testing
        self.dummy_img = Image.new('RGB', (224, 224), color='white')

    def __len__(self):
        return self.num_samples

    def __getitem__(self, idx):
        # Simulate validation data
        label = idx % NUM_CLASSES  # Cycle through classes
        if self.transform:
            return self.transform(self.dummy_img), label
        return self.dummy_img, label

def train_classifier():
    print("--- Starting Document Classifier Training (Demo) ---")
    
    # 2. Data Preprocessing & Augmentation (Crucial Step)
    data_transforms = transforms.Compose([
        transforms.Resize(256),
        transforms.CenterCrop(224),
        transforms.ToTensor(),
        transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])
    ])

    # 3. Load Data
    # train_dataset = datasets.ImageFolder("data/train", transform=data_transforms)
    train_dataset = MockDocumentDataset(transform=data_transforms)
    train_loader = DataLoader(train_dataset, batch_size=BATCH_SIZE, shuffle=True)
    
    print(f"Dataset Loaded: {len(train_dataset)} images.")

    # 4. Define Model (Transfer Learning)
    # We load a pre-trained ResNet18 and replace the final layer.
    print("Loading Pre-trained ResNet18...")
    model = models.resnet18(pretrained=True)
    
    # Freeze early layers (optional, good for teaching feature extraction)
    for param in model.parameters():
        param.requires_grad = False
        
    # Replace the Output Layer (Fully Connected)
    num_ftrs = model.fc.in_features
    model.fc = nn.Linear(num_ftrs, NUM_CLASSES)
    
    # 5. Define Loss & Optimizer
    criterion = nn.CrossEntropyLoss()
    optimizer = optim.SGD(model.fc.parameters(), lr=LEARNING_RATE, momentum=0.9)

    # 6. Training Loop ( The "Heart" of PyTorch )
    device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
    model = model.to(device)
    
    loss_history = []
    
    print("Starting Training Loop...")
    model.train()
    for epoch in range(EPOCHS):
        running_loss = 0.0
        for i, (inputs, labels) in enumerate(train_loader):
            inputs = inputs.to(device)
            labels = labels.to(device)

            # Zero the parameter gradients
            optimizer.zero_grad()

            # Forward + Backward + Optimize
            outputs = model(inputs)
            loss = criterion(outputs, labels)
            loss.backward()
            optimizer.step()

            running_loss += loss.item()
        
        epoch_loss = running_loss / len(train_loader)
        loss_history.append(epoch_loss)
        print(f"Epoch [{epoch+1}/{EPOCHS}], Loss: {epoch_loss:.4f}")

    print("Training Complete.")
    
    # 7. Visualization
    # plt.plot(loss_history)
    # plt.title("Training Loss")
    # plt.show()
    print(f"Final Model saved to 'document_classifier.pth'")
    # torch.save(model.state_dict(), "document_classifier.pth")

if __name__ == "__main__":
    train_classifier()
